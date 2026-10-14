import pandas as pd
from docxtpl import DocxTemplate, RichText
import os
import time
from docx2pdf import convert
import fitz 

def formatar_nome_aluno(nome):
    nome_str = str(nome).strip()
    if nome_str.lower() == 'nan' or not nome_str:
        return ""
        
    excecoes = ['da', 'de', 'do', 'das', 'dos', 'e']
    palavras = nome_str.split()
    nome_formatado = []
    for p in palavras:
        if p.lower() in excecoes:
            nome_formatado.append(p.lower())
        else:
            nome_formatado.append(p.capitalize())
            
    return ' '.join(nome_formatado)

def formatar_professor(nome):
    nome_str = str(nome).strip()
    if nome_str.lower() == 'nan' or not nome_str:
        return ""
        
    excecoes = ['da', 'de', 'do', 'das', 'dos', 'e']
    palavras = nome_str.split()
    nome_formatado = []
    
    for p in palavras:
        if p.lower() in excecoes:
            nome_formatado.append(p.lower())
        else:
            nome_formatado.append(p.capitalize())
            
    nome_final = ' '.join(nome_formatado)
    
    if not palavras:
        return ""
        
    primeiro_nome = palavras[0].lower()
    femininos_comuns = ['beatriz', 'raquel', 'carmen', 'simone', 'aline', 'viviane', 'gisele', 'cleide', 'sueli', 'iris', 'elis', 'jane', 'ivone', 'ester', 'thais', 'thaís', 'taís', 'daniele', 'tatiane', 'fabiany']
    
    if primeiro_nome.endswith('a') or primeiro_nome in femininos_comuns:
        return f"Profª. Drª. {nome_final}"
    else:
        return f"Prof. Dr. {nome_final}"

df_dados = pd.read_excel("dados_qualificação_mestrado.xlsx", dtype=str)

for index, linha in df_dados.iterrows():
    aluno_texto = formatar_nome_aluno(linha.get('Aluno', ''))
    if not aluno_texto:
        continue
        
    print(f"\n--- Processando aluno(a): {aluno_texto} ---")
        
    titulo_texto = str(linha.get('Título', '')).strip()
    if titulo_texto.lower() == 'nan':
        titulo_texto = ""
        
    tamanho_titulo = len(titulo_texto)
    if tamanho_titulo <= 20:
        fonte_titulo = 36  # Fonte 18
    elif tamanho_titulo <= 125:
        fonte_titulo = 32  # Fonte 16
    elif tamanho_titulo <= 160:
        fonte_titulo = 28  # Fonte 14
    else:
        fonte_titulo = 24  # Fonte 12
    titulo_rt = RichText(titulo_texto, size=fonte_titulo)
        
    data_bruta = str(linha.get('Data', '')).strip().split()[0]  
    if '-' in data_bruta:
        partes = data_bruta.split('-')
        data_formatada = f"{partes[2]}/{partes[1]}/{partes[0]}" if len(partes) == 3 else data_bruta
    else:
        data_formatada = data_bruta
    if data_formatada.lower() == 'nan': data_formatada = ""

    hora_bruta = str(linha.get('Hora', '')).strip()
    hora_formatada = "" if hora_bruta.lower() == 'nan' else (hora_bruta[:5] if len(hora_bruta) >= 5 else hora_bruta)
        
    sala_bruta = str(linha.get('Sala', '')).replace(".0", "").strip()
    sala_formatada = "" if sala_bruta.lower() == 'nan' else sala_bruta
    
    uni_externo = str(linha.get('Universidade Externo', '')).strip()
    uni_externo = "" if uni_externo.lower() == 'nan' else uni_externo
    
    uni_suplente = str(linha.get('Universidade Suplente', '')).strip()
    uni_suplente = "" if uni_suplente.lower() == 'nan' else uni_suplente

    membro_externo1_texto = formatar_professor(linha.get('Membro Externo', ''))

    contexto = {
        'titulo': titulo_rt, 
        'titulo_convite': titulo_texto,
        'aluno': aluno_texto,
        'orientador': formatar_professor(linha.get('Orientador', '')),
        'presidente': formatar_professor(linha.get('Presidente', '')),
        'membro_interno': formatar_professor(linha.get('Membro Interno', '')),
        'suplente': formatar_professor(linha.get('Suplente', '')),
        'membro_externo': membro_externo1_texto, 
        'universidade_externo': uni_externo,
        'universidade_suplente': uni_suplente,
        'data': data_formatada,
        'hora': hora_formatada,
        'numero_sala': sala_formatada
    }
    
    nome_da_pasta = aluno_texto
    os.makedirs(nome_da_pasta, exist_ok=True)
    nome_minusculo = aluno_texto.lower()
    
    try:
        print(" > Gerando Cartaz de Defesa...")
        doc_defesa = DocxTemplate("modelo_qualificação_mestrado.docx")
        doc_defesa.render(contexto)
        
        caminho_docx_def = os.path.abspath(os.path.join(nome_da_pasta, f"Cartaz - {nome_minusculo}.docx"))
        caminho_pdf_def = os.path.abspath(os.path.join(nome_da_pasta, f"Cartaz - {nome_minusculo}.pdf"))
        caminho_png_def = os.path.abspath(os.path.join(nome_da_pasta, f"Cartaz - {nome_minusculo}.png"))
        
        doc_defesa.save(caminho_docx_def)
        convert(caminho_docx_def, caminho_pdf_def)
        
        if os.path.exists(caminho_pdf_def):
            pdf_documento = fitz.open(caminho_pdf_def)
            pagina = pdf_documento.load_page(0)
            pix = pagina.get_pixmap(dpi=300)
            pix.save(caminho_png_def)
            pdf_documento.close()
            os.remove(caminho_pdf_def)
            
    except Exception as e:
        print(f" [ERRO] Falha ao gerar cartaz: {e}")

    try:
        print(" > Gerando Convite da Banca...")
        doc_convite = DocxTemplate("convite_qualificação_mestrado.docx")
        doc_convite.render(contexto)
        
        caminho_docx_conv = os.path.abspath(os.path.join(nome_da_pasta, f"Convite - {nome_minusculo}.docx"))
        caminho_pdf_conv = os.path.abspath(os.path.join(nome_da_pasta, f"Convite - {nome_minusculo}.pdf"))
        caminho_png_conv = os.path.abspath(os.path.join(nome_da_pasta, f"Convite - {nome_minusculo}.png"))
        
        doc_convite.save(caminho_docx_conv)
        convert(caminho_docx_conv, caminho_pdf_conv)
        
        if os.path.exists(caminho_pdf_conv):
            pdf_convite_doc = fitz.open(caminho_pdf_conv)
            pagina_conv = pdf_convite_doc.load_page(0)
            pix_conv = pagina_conv.get_pixmap(dpi=300)
            pix_conv.save(caminho_png_conv)
            pdf_convite_doc.close()
            os.remove(caminho_pdf_conv)
            
    except Exception as e:
        print(f" [ERRO] Falha ao gerar convite: {e}")

    time.sleep(1.5)

print("\nAutomação concluída com sucesso!")