import sys
import subprocess

def auto_instalar():
    pacotes = {
        "docxtpl": "docxtpl",
        "docx2pdf": "docx2pdf"
    }
    for modulo, pacote in pacotes.items():
        try:
            __import__(modulo)
        except ImportError:
            subprocess.check_call([sys.executable, "-m", "pip", "install", pacote])

auto_instalar()

from docxtpl import DocxTemplate
from datetime import datetime
import locale
import os
import time
from docx2pdf import convert

try:
    locale.setlocale(locale.LC_TIME, 'pt_BR.utf8')
except locale.Error:
    pass 

nome = input("Digite o nome: ").strip()
matricula = input("Digite a matrícula: ").strip()
titulo = input("Digite o título da dissertação: ").strip()

data_hoje = datetime.now().strftime("%d de %B de %Y")

doc = DocxTemplate("Modelo_declaração_conclusão.docx")

contexto = {
    'nome': nome,
    'matricula': matricula,
    'titulo': titulo,
    'dia_hoje': data_hoje
}

doc.render(contexto)

nome_da_pasta = nome
os.makedirs(nome_da_pasta, exist_ok=True)

nome_minusculo = nome.lower()
nome_arquivo_docx = f"Declaração - {nome_minusculo}.docx"
nome_arquivo_pdf = f"Declaração - {nome_minusculo}.pdf"

caminho_docx = os.path.join(nome_da_pasta, nome_arquivo_docx)
caminho_pdf = os.path.join(nome_da_pasta, nome_arquivo_pdf)

doc.save(caminho_docx)

try:
    convert(caminho_docx, caminho_pdf)
except Exception:
    pass

if os.path.exists(caminho_pdf):
    os.remove(caminho_docx)
    
time.sleep(1.5)