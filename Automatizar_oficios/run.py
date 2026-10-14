import sys
import subprocess

def auto_instalar():
    pacotes = {
        "pandas": "pandas",
        "docxtpl": "docxtpl",
        "docx2pdf": "docx2pdf",
        "openpyxl": "openpyxl"
    }
    for modulo, pacote in pacotes.items():
        try:
            __import__(modulo)
        except ImportError:
            subprocess.check_call([sys.executable, "-m", "pip", "install", pacote])

auto_instalar()

import pandas as pd
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

data_hoje = datetime.now().strftime("%d de %B de %Y")
ano_atual = datetime.now().year

print("=== GERADOR DE OFÍCIOS ===")
entrada_oficio = input("Digite o número do PRIMEIRO Ofício: ").strip()
numero_oficio_atual = int(entrada_oficio) if entrada_oficio.isdigit() else 1

df_alunos = pd.read_excel("dados_alunos.xlsx")

print(f"Planilha carregada. Foram encontrados {len(df_alunos)} alunos.")
print("-" * 30)

for index, linha in df_alunos.iterrows():
    
    doc = DocxTemplate("modelo_oficio.docx")
    
    nome = str(linha['Nome']).strip()
    matricula = str(linha['Matricula']).strip()
    
    texto_oficio = f"{numero_oficio_atual}/{ano_atual}"
    
    contexto = {
        'numero_oficio': texto_oficio,
        'data_atual': data_hoje,
        'nome_aluno': nome,
        'matricula_aluno': matricula
    }
    
    doc.render(contexto)
    
    nome_da_pasta = nome
    os.makedirs(nome_da_pasta, exist_ok=True)
    
    nome_minusculo = nome.lower()
    nome_arquivo_docx = f"0 - oficio - {nome_minusculo}.docx"
    nome_arquivo_pdf = f"0 - oficio - {nome_minusculo}.pdf"
    
    caminho_docx = os.path.join(nome_da_pasta, nome_arquivo_docx)
    caminho_pdf = os.path.join(nome_da_pasta, nome_arquivo_pdf)
    
    doc.save(caminho_docx)
    print(f"Processando: {nome}...")
    
    try:
        convert(caminho_docx, caminho_pdf)
    except Exception as e:
        print(f"  -> Aviso ignorado ao fechar o Word: {e}")
    
    if os.path.exists(caminho_pdf):
        os.remove(caminho_docx)
        print(f"  -> Sucesso! PDF salvo em '{nome_da_pasta}'\n")
    else:
        print(f"  -> Erro: O PDF de {nome} não foi gerado.\n")
    
    numero_oficio_atual += 1
    
    time.sleep(1.5) 

print("-" * 30)
print("Automação concluída!")