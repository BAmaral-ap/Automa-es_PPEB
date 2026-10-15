"""Caixas fixas, composição do texto e conferência do PDF renderizado."""
from dataclasses import dataclass
from functools import lru_cache
import math
import re
import unicodedata

from lxml import etree
import pymupdf

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "v": "urn:schemas-microsoft-com:vml",
}


@dataclass(frozen=True)
class Caixa:
    x: float
    y: float
    largura: float
    altura: float
    fonte: float
    centro: bool = True


# Medidas em pontos, relativas à página A4 e aos rótulos da imagem original.
CAIXAS = {
    "titulo": Caixa(82, 229, 498, 76, 21, False),
    "aluno": Caixa(98, 320, 482, 34, 19),
    "orientador": Caixa(142, 360, 438, 32, 17),
    "banca": Caixa(12, 425, 576, 238, 18, False),
    "data": Caixa(92, 680, 186, 34, 20),
    "hora": Caixa(398, 680, 186, 34, 20),
    "sala": Caixa(92, 720, 492, 30, 20),
}
MARGEM_X, MARGEM_Y = 7, 4
FONTE_MINIMA = 8
METRICA = pymupdf.Font("hebo")


@lru_cache(maxsize=8192)
def largura_texto(texto, fonte):
    return METRICA.text_length(texto, fontsize=fonte)


@dataclass
class Composicao:
    fonte: float
    paragrafos: list[list[str]]
    espaco_paragrafos: float = 0


def normalizar(texto):
    return " ".join(texto.split())


def quebrar_linhas(texto, largura, fonte):
    """Mede as palavras; também divide palavras sem espaços que não caibam."""
    linhas, atual = [], ""
    for palavra in normalizar(texto).split():
        candidato = f"{atual} {palavra}" if atual else palavra
        if largura_texto(candidato, fonte) <= largura:
            atual = candidato
            continue
        if atual:
            linhas.append(atual)
            atual = ""
        if largura_texto(palavra, fonte) <= largura:
            atual = palavra
            continue
        for letra in palavra:
            if atual and largura_texto(atual + letra, fonte) > largura:
                linhas.append(atual)
                atual = ""
            atual += letra
    if atual:
        linhas.append(atual)
    return linhas or [""]


def compor(chave, textos, limite=None):
    caixa = CAIXAS[chave]
    maximo = caixa.fonte if limite is None else limite
    # Calibri/Carlito e Helvetica têm métricas diferentes. A margem de 6% evita
    # que o Word volte a quebrar uma linha que já foi composta pelo script.
    largura = (caixa.largura - 2 * MARGEM_X) * 0.94
    altura = caixa.altura - 2 * MARGEM_Y - 2
    def tentar(fonte):
        paragrafos = [quebrar_linhas(t, largura, fonte) for t in textos if t.strip()]
        altura_linhas = sum(len(p) for p in paragrafos) * fonte * 1.2
        intervalos = max(0, len(paragrafos) - 1) if chave == "banca" else 0
        altura_texto = altura_linhas + intervalos * 5
        if altura_texto <= altura:
            # Distribui o espaço livre entre os membros, em vez de deixá-lo
            # acumulado abaixo do último. Arredonda para baixo em twips.
            espaco = math.floor((altura - altura_linhas) / intervalos * 20) / 20 if intervalos else 0
            return Composicao(fonte, paragrafos, espaco)
        return None

    # Rejeita textos impossíveis antes de medir todos os tamanhos intermediários.
    if tentar(FONTE_MINIMA) is not None:
        for meio_ponto in range(math.floor(maximo * 2), FONTE_MINIMA * 2 - 1, -1):
            composicao = tentar(meio_ponto / 2)
            if composicao is not None:
                return composicao
    raise ValueError(f"O texto de '{chave}' não cabe na caixa nem com fonte "
                     f"{FONTE_MINIMA}. Reduza o texto; a caixa não será ampliada.")


def textos_caixas(contexto):
    banca = [contexto[chave] for chave in (
        "presidente", "membro_interno", "membro_externo_1", "membro_externo_2",
        "membro_externo_3", "suplente_1", "suplente_2") if contexto[chave]]
    return {chave: ([f"• {t}" for t in banca] if chave == "banca" else [contexto[chave]])
            for chave in CAIXAS}


def elemento(tag, **attrs):
    prefixo, local = tag.split(":")
    el = etree.Element(f"{{{NS[prefixo]}}}{local}")
    for nome, valor in attrs.items():
        el.set(f"{{{NS['w']}}}{nome}", str(valor))
    return el


def paragrafos_xml(chave, composicao):
    saida = []
    for i, linhas in enumerate(composicao.paragrafos):
        p = elemento("w:p")
        props = elemento("w:pPr")
        props.append(elemento("w:spacing", before=0,
                              after=round(composicao.espaco_paragrafos * 20) if i < len(composicao.paragrafos) - 1 else 0,
                              line=round(composicao.fonte * 1.2 * 20), lineRule="exact"))
        props.append(elemento("w:jc", val="center" if CAIXAS[chave].centro else "both"))
        props.append(elemento("w:widowControl", val=0))
        p.append(props)
        r = elemento("w:r")
        rprops = elemento("w:rPr")
        rprops.append(elemento("w:rFonts", ascii="Calibri", hAnsi="Calibri", cs="Calibri"))
        rprops.append(elemento("w:b"))
        for tag in ("sz", "szCs"):
            rprops.append(elemento("w:" + tag, val=round(composicao.fonte * 2)))
        r.append(rprops)
        for j, linha in enumerate(linhas):
            if j:
                r.append(elemento("w:br"))
            t = elemento("w:t")
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t.text = linha
            r.append(t)
        p.append(r)
        saida.append(p)
    return saida


def aplicar_composicao(xml, composicoes):
    for chave, composicao in composicoes.items():
        anchors = xml.xpath(f'//wp:anchor[wp:docPr/@name="PPEB_{chave}"]', namespaces=NS)
        shapes = xml.xpath(f'//v:shape[@id="PPEB_{chave}"]', namespaces=NS)
        if len(anchors) != 1 or len(shapes) != 1:
            raise ValueError(f"Modelo incompatível: caixa fixa '{chave}' não encontrada. "
                             "Use o modelo atualizado junto com o script.")
        for container in (anchors[0], shapes[0]):
            contents = container.xpath(".//w:txbxContent", namespaces=NS)
            if len(contents) != 1:
                raise ValueError(f"Caixa '{chave}' inválida no modelo.")
            content = contents[0]
            content.clear()
            content.extend(paragrafos_xml(chave, composicao))


def canonico(texto):
    texto = unicodedata.normalize("NFKC", texto).replace("\u00ad", "")
    return re.sub(r"\s+", "", texto)


def conferir_pdf(pdf, contexto):
    """Confere o conteúdo visível e os limites de cada caixa, não só as páginas."""
    esperados = textos_caixas(contexto)
    problemas = []
    with pymupdf.open(pdf) as doc:
        if len(doc) != 1:
            return list(CAIXAS)
        pagina = doc[0]
        for chave, caixa in CAIXAS.items():
            rect = pymupdf.Rect(caixa.x, caixa.y, caixa.x + caixa.largura, caixa.y + caixa.altura)
            visivel = pagina.get_text("text", clip=rect, sort=True)
            if canonico(visivel) != canonico(" ".join(esperados[chave])):
                problemas.append(chave)
                continue
            # Aceita apenas a tolerância de arredondamento da exportação.
            limite = rect + (-0.75, -0.75, 0.75, 0.75)
            for palavra in pagina.get_text("words", clip=rect):
                if not limite.contains(pymupdf.Rect(palavra[:4])):
                    problemas.append(chave)
                    break
    return problemas
