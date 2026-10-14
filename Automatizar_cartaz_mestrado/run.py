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
    femininos_comuns = ['beatriz', 'raquel', 'carmen', 'simone', 'aline', 'viviane', 'gisele', 'cleide', 'sueli', 'iris', 'elis', 'jane', 'ivone', 'ester', 'thais', 'thaís', 'taís']
    
    if primeiro_nome.endswith('a') or primeiro_nome in femininos_comuns:
        return f"Profª. Drª. {nome_final}"
    else:
        return f"Prof. Dr. {nome_final}"

def get_genero(nome):
    nome_str = str(nome).strip()
    if not nome_str or nome_str.lower() == 'nan':
        return 'M'
        
    palavras = nome_str.split()
    if not palavras:
        return 'M'
        
    primeiro_nome = palavras[0].lower()
    femininos_comuns = ['beatriz', 'raquel', 'carmen', 'simone', 'aline', 'viviane', 'gisele', 'cleide', 'sueli', 'iris', 'elis', 'jane', 'ivone', 'ester', 'thais', 'thaís', 'taís']
    
    if primeiro_nome.endswith('a') or primeiro_nome in femininos_comuns:
        return 'F'
    return 'M'

def datas_por_extenso(data_str):
    meses = {'01':'janeiro', '02':'fevereiro', '03':'março', '04':'abril', '05':'maio', '06':'junho', '07':'julho', '08':'agosto', '09':'setembro', '10':'outubro', '11':'novembro', '12':'dezembro'}
    dias = {'01':'primeiro', '02':'dois', '03':'três', '04':'quatro', '05':'cinco', '06':'seis', '07':'sete', '08':'oito', '09':'nove', '10':'dez', '11':'onze', '12':'doze', '13':'treze', '14':'quatorze', '15':'quinze', '16':'dezesseis', '17':'dezessete', '18':'dezoito', '19':'dezenove', '20':'vinte', '21':'vinte e um', '22':'vinte e dois', '23':'vinte e três', '24':'vinte e quatro', '25':'vinte e cinco', '26':'vinte e seis', '27':'vinte e sete', '28':'vinte e oito', '29':'vinte e nove', '30':'trinta', '31':'trinta e um'}
    
    if '/' in data_str:
        d, m, a = data_str.split('/')
        dia_ext = dias.get(d, d)
        mes_ext = meses.get(m, m)
        texto_ata = f"{dia_ext} dias do mês de {mes_ext} de {a}"
        texto_cidade = f"{int(d)} de {mes_ext} de {a}"
        return texto_ata, texto_cidade
    return "", ""

print("=== GERADOR DE DOCUMENTOS DE DEFESA ===")
entrada_ata = input("Digite o número da PRIMEIRA Ata (ou pressione ENTER para deixar em branco): ").strip()
numero_ata_atual = int(entrada_ata) if entrada_ata.isdigit() else None

df_dados = pd.read_excel("dados_defesa.xlsx", dtype=str)

for index, linha in df_dados.iterrows():
    aluno_texto = formatar_nome_aluno(linha.get('Aluno', ''))
    if not aluno_texto:
        continue
        
    print(f"\n--- Processando aluno(a): {aluno_texto} ---")
        
    titulo_texto = str(linha.get('Título', '')).strip()
    if titulo_texto.lower() == 'nan':
        titulo_texto = ""
        
    tamanho_titulo = len(titulo_texto)
    if tamanho_titulo <= 90:
        fonte_titulo = 36  
    elif tamanho_titulo <= 125:
        fonte_titulo = 32  
    elif tamanho_titulo <= 160:
        fonte_titulo = 28  
    else:
        fonte_titulo = 24  
        
    titulo_rt = RichText(titulo_texto, size=fonte_titulo)
        
    data_bruta = str(linha.get('Data', '')).strip().split()[0]  
    if '-' in data_bruta:
        partes = data_bruta.split('-')
        data_formatada = f"{partes[2]}/{partes[1]}/{partes[0]}" if len(partes) == 3 else data_bruta
    else:
        data_formatada = data_bruta
    if data_formatada.lower() == 'nan': data_formatada = ""

    data_extenso, data_cidade = datas_por_extenso(data_formatada)

    hora_bruta = str(linha.get('Hora', '')).strip()
    hora_formatada = "" if hora_bruta.lower() == 'nan' else (hora_bruta[:5] if len(hora_bruta) >= 5 else hora_bruta)
        
    sala_bruta = str(linha.get('Sala', '')).replace(".0", "").strip()
    sala_formatada = "" if sala_bruta.lower() == 'nan' else sala_bruta
    
    uni_externo = str(linha.get('Universidade Externo', '')).strip()
    uni_externo = "" if uni_externo.lower() == 'nan' else uni_externo
    
    uni_suplente = str(linha.get('Universidade Suplente', '')).strip()
    uni_suplente = "" if uni_suplente.lower() == 'nan' else uni_suplente

    membro_externo1_texto = formatar_professor(linha.get('Membro Externo', ''))

    if numero_ata_atual is not None:
        numero_ata_formatado = str(numero_ata_atual)
        numero_ata_atual += 1
    else:
        numero_ata_formatado = "_______"

    gen_ori = get_genero(linha.get('Orientador', ''))
    desc_ori = "Orientadora do trabalho" if gen_ori == 'F' else "Orientador do trabalho"

    gen_int = get_genero(linha.get('Membro Interno', ''))
    desc_int = "Examinadora interna" if gen_int == 'F' else "Examinador interno"

    gen_ext = get_genero(linha.get('Membro Externo', ''))
    desc_ext = "Examinadora externa" if gen_ext == 'F' else "Examinador externo"
    
    orientador_bruto = str(linha.get('Orientador', '')).strip()
    presidente_bruto = str(linha.get('Presidente', '')).strip()
    if presidente_bruto.lower() == 'nan' or not presidente_bruto:
        presidente_bruto = orientador_bruto

    contexto_base = {
        'titulo_convite': titulo_texto,
        'titulo': titulo_texto,
        'Título': titulo_texto,
        'aluno': aluno_texto,
        'aluno_maiusculo': aluno_texto.upper(),
        'orientador': formatar_professor(orientador_bruto),
        'presidente': formatar_professor(presidente_bruto),
        'membro_interno': formatar_professor(linha.get('Membro Interno', '')),
        'suplente': formatar_professor(linha.get('Suplente', '')),
        'membro_externo': membro_externo1_texto, 
        'universidade_externo': uni_externo,
        'universidade_suplente': uni_suplente,
        'desc_orientador': desc_ori,
        'desc_interno': desc_int,
        'desc_externo': desc_ext,
        'data': data_formatada,
        'data_extenso': data_extenso,
        'dia': data_cidade,
        'hora': hora_formatada,
        'numero_sala': sala_formatada,
        'numero_ata': numero_ata_formatado
    }
    
    contexto_cartaz = contexto_base.copy()
    contexto_cartaz['titulo'] = titulo_rt
    
    nome_da_pasta = aluno_texto
    os.makedirs(nome_da_pasta, exist_ok=True)
    nome_minusculo = aluno_texto.lower()
    
    try:
        print(" > Gerando Cartaz de Defesa...")
        doc_defesa = DocxTemplate("modelo_defesa.docx")
        doc_defesa.render(contexto_cartaz)
        
        caminho_docx_def = os.path.abspath(os.path.join(nome_da_pasta, f"0 - cartaz - {nome_minusculo}.docx"))
        caminho_pdf_def = os.path.abspath(os.path.join(nome_da_pasta, f"0 - cartaz - {nome_minusculo}.pdf"))
        caminho_png_def = os.path.abspath(os.path.join(nome_da_pasta, f"0 - cartaz - {nome_minusculo}.png"))
        
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
        doc_convite = DocxTemplate("modelo_convite.docx")
        doc_convite.render(contexto_base)
        
        caminho_docx_conv = os.path.abspath(os.path.join(nome_da_pasta, f"1 - convite - {nome_minusculo}.docx"))
        caminho_pdf_conv = os.path.abspath(os.path.join(nome_da_pasta, f"1 - convite - {nome_minusculo}.pdf"))
        caminho_png_conv = os.path.abspath(os.path.join(nome_da_pasta, f"1 - convite - {nome_minusculo}.png"))
        
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

    try:
        print(" > Gerando Ata de Defesa...")
        doc_ata = DocxTemplate("ata_defesa_mestrado.docx")
        doc_ata.render(contexto_base)
        
        caminho_docx_ata = os.path.abspath(os.path.join(nome_da_pasta, f"2 - ata - {nome_minusculo}.docx"))
        caminho_pdf_ata = os.path.abspath(os.path.join(nome_da_pasta, f"2 - ata - {nome_minusculo}.pdf"))
        
        doc_ata.save(caminho_docx_ata)
        convert(caminho_docx_ata, caminho_pdf_ata)
            
    except Exception as e:
        print(f" [ERRO] Falha ao gerar ata: {e}")

    time.sleep(1.5)

print("\nAutomação concluída com sucesso!")