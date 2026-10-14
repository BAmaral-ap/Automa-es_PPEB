import os
from docx2pdf import convert
import fitz 

print("=== CONVERSOR MANUAL DE DOCX PARA PNG ===")
print("AVISO: Feche o arquivo no Word antes de rodar este código para evitar erros.\n")

nome_arquivo = input("Digite o nome do arquivo (sem o .docx): ").strip()

caminho_docx = os.path.abspath(f"{nome_arquivo}.docx")
caminho_pdf = os.path.abspath(f"{nome_arquivo}.pdf")
caminho_png = os.path.abspath(f"{nome_arquivo}.png")

if not os.path.exists(caminho_docx):
    print(f"\n[ERRO] O arquivo '{nome_arquivo}.docx' não foi encontrado nesta pasta.")
else:
    try:
        print(" > Convertendo para PDF...")
        convert(caminho_docx, caminho_pdf)
        
        if os.path.exists(caminho_pdf):
            print(" > Convertendo para PNG...")
            pdf_documento = fitz.open(caminho_pdf)
            pagina = pdf_documento.load_page(0)
            pix = pagina.get_pixmap(dpi=300)
            pix.save(caminho_png)
            pdf_documento.close()
            
            os.remove(caminho_pdf)
            print(f"\nSucesso! Imagem '{nome_arquivo}.png' gerada.")
            
    except Exception as e:
        print(f"\n[ERRO] Falha na conversão: {e}")

input("\nPressione ENTER para fechar...")