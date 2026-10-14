from contextlib import redirect_stdout, redirect_stderr
from datetime import datetime, time
import io
from pathlib import Path
import sys
import tempfile
import unittest
from zipfile import ZipFile

from lxml import etree
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import run


class AutomacaoTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.planilha = self.root / "dados.xlsx"
        original = load_workbook(run.BASE / "dados_qualificacao_doutorado.xlsx")
        original.save(self.planilha)
        original.close()

    def alterar(self, coluna, valor):
        wb = load_workbook(self.planilha)
        ws = wb["Dados"]
        headers = [c.value for c in ws[1]]
        ws.cell(2, headers.index(coluna) + 1, valor)
        wb.save(self.planilha)
        wb.close()

    def executar(self, *extras):
        with redirect_stdout(io.StringIO()), redirect_stderr(io.StringIO()):
            return run.main(["--planilha", str(self.planilha), "--saida", str(self.root / "saida"),
                             "--formato", "docx", *extras])

    def test_geracao_preserva_imagens_e_estrutura(self):
        self.assertEqual(self.executar(), 0)
        docx = next((self.root / "saida").glob("*/*.docx"))
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        with ZipFile(docx) as rendered, ZipFile(run.BASE / "modelo_qualificacao_doutorado.docx") as template:
            self.assertIsNone(rendered.testzip())
            xml = rendered.read("word/document.xml")
            tree = etree.fromstring(xml)
            textos = tree.xpath("//w:t/text()", namespaces=ns)
            self.assertIn("ADRIAN SOUZA DOS SANTOS", textos)
            self.assertIn("31/03/2026", textos)
            self.assertIn("14h00", textos)
            self.assertNotIn(b"{{", xml)
            self.assertEqual(sum(" - Membro Externo" in t for t in textos), 3)
            self.assertEqual(sum(" - Suplente" in t for t in textos), 2)
            media = [n for n in template.namelist() if n.startswith("word/media/")]
            self.assertTrue(media)
            for nome in media:
                self.assertEqual(template.read(nome), rendered.read(nome))
            before = etree.fromstring(template.read("word/document.xml"))
            self.assertEqual(len(before.xpath("//w:p", namespaces=ns)), len(tree.xpath("//w:p", namespaces=ns)))

    def test_caracteres_especiais_nao_quebram_xml(self):
        titulo = "Educação & sociedade: <escola> {{ texto literal }}"
        self.alterar("Título", titulo)
        self.assertEqual(self.executar(), 0)
        with ZipFile(next((self.root / "saida").glob("*/*.docx"))) as z:
            tree = etree.fromstring(z.read("word/document.xml"))
            self.assertIn(titulo, tree.xpath('//*[local-name()="t"]/text()'))

    def test_campos_opcionais_nao_deixam_rotulos(self):
        wb = load_workbook(self.planilha)
        ws = wb["Dados"]
        for cell in ws[1]:
            if cell.value in ("Membro Externo 3", "Universidade Externo 3", "Suplente 2", "Universidade Suplente 2"):
                ws.cell(2, cell.column).value = None
        wb.save(self.planilha)
        wb.close()
        self.assertEqual(self.executar(), 0)
        with ZipFile(next((self.root / "saida").glob("*/*.docx"))) as z:
            tree = etree.fromstring(z.read("word/document.xml"))
            texts = tree.xpath('//*[local-name()="t"]/text()')
            self.assertEqual(sum(" - Membro Externo" in t for t in texts), 2)
            self.assertEqual(sum(" - Suplente" in t for t in texts), 1)
            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            self.assertEqual(len(tree.xpath("//w:pPr/w:numPr", namespaces=ns)), 5)

    def test_nao_sobrescreve_sem_opcao(self):
        self.assertEqual(self.executar(), 0)
        path = next((self.root / "saida").glob("*/*.docx"))
        before = path.read_bytes()
        self.alterar("Título", "Título novo")
        self.assertEqual(self.executar(), 1)
        self.assertEqual(path.read_bytes(), before)
        self.assertEqual(self.executar("--sobrescrever"), 0)
        self.assertNotEqual(path.read_bytes(), before)

    def test_entrada_invalida_nao_produz_arquivos(self):
        self.alterar("Data", "31/02/2026")
        self.assertEqual(self.executar(), 1)
        self.assertFalse((self.root / "saida").exists())

    def test_formula_rejeitada(self):
        self.alterar("Título", '=CONCAT("a","b")')
        self.assertEqual(self.executar(), 1)

    def test_aluno_duplicado_rejeitado(self):
        wb = load_workbook(self.planilha)
        ws = wb["Dados"]
        ws.append([c.value for c in ws[2]])
        wb.save(self.planilha)
        wb.close()
        self.assertEqual(self.executar(), 1)
        self.assertFalse((self.root / "saida").exists())

    def test_formatacao_de_datas_horas_excel(self):
        self.assertEqual(run.formatar_data(datetime(2026, 3, 31)), "31/03/2026")
        self.assertEqual(run.formatar_hora(time(14)), "14h00")
        self.assertEqual(run.formatar_hora(14 / 24), "14h00")
        self.assertEqual(run.formatar_hora("9:05"), "09h05")
        with self.assertRaises(ValueError):
            run.formatar_hora("24h00")

    def test_nome_pasta_compativel_com_windows(self):
        self.assertEqual(run.nome_pasta("../Ana/João"), "_Ana_João")
        self.assertEqual(run.nome_pasta("CON"), "Aluno - CON")

    def test_pdf_multipagina_nao_vira_imagem_incompleta(self):
        import pymupdf
        path = self.root / "duas_paginas.pdf"
        with pymupdf.open() as pdf:
            pdf.new_page()
            pdf.new_page()
            pdf.save(path)
        with self.assertRaisesRegex(ValueError, "2 páginas"):
            run.gerar_png(path, 72)
        self.assertFalse(path.with_suffix(".png").exists())

    def test_falha_de_conversor_retorna_erro_e_nao_publica_docx(self):
        from unittest.mock import patch
        with patch.object(run, "escolher_conversor", return_value="libreoffice"), \
             patch.object(run, "gerar_pdf", side_effect=RuntimeError("Falha simulada")), \
             redirect_stdout(io.StringIO()), redirect_stderr(io.StringIO()):
            code = run.main(["--planilha", str(self.planilha), "--saida", str(self.root / "saida")])
        self.assertEqual(code, 1)
        self.assertFalse(list((self.root / "saida").glob("*/*.docx")))


if __name__ == "__main__":
    unittest.main()
