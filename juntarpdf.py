import os
import glob
from PyPDF2 import PdfMerger # Necessário: pip install pypdf2

# --- Configurações ---
SAIDA_PASTA = 'Documentos_Gerados'
NOME_FINAL_PDF = 'Catalogo_Final_UNIDO.pdf'

def unir_arquivos_pdf():
    """
    Escaneia a pasta de saída, encontra todos os PDFs gerados e os une 
    em um único arquivo final.
    """
    
    # 1. Verificar se a pasta existe
    if not os.path.exists(SAIDA_PASTA):
        print(f"ERRO: A pasta de saída '{SAIDA_PASTA}' não foi encontrada.")
        print("Execute o script de automação de documentos primeiro.")
        return
    
    # 2. Encontrar todos os PDFs de Lote
    # Filtra todos os PDFs que não são o arquivo final
    padrao_busca = os.path.join(SAIDA_PASTA, 'Catalogo_Lote_*.pdf')
    arquivos_pdf_gerados = sorted(glob.glob(padrao_busca)) # Ordena para garantir a sequência correta
    
    if not arquivos_pdf_gerados:
        print(f"AVISO: Nenhum arquivo PDF de lote foi encontrado na pasta '{SAIDA_PASTA}'.")
        return

    caminho_final_pdf = os.path.join(SAIDA_PASTA, NOME_FINAL_PDF)
    
    # 3. Iniciar o Merge
    print(f"\n--- INICIANDO MERGE DE {len(arquivos_pdf_gerados)} PDFS ---")
    try:
        merger = PdfMerger()
        
        for pdf_file in arquivos_pdf_gerados:
            nome_base = os.path.basename(pdf_file)
            print(f"  -> Adicionando {nome_base}")
            merger.append(pdf_file)

        # 4. Salvar o arquivo final
        with open(caminho_final_pdf, "wb") as saida:
            merger.write(saida)
        
        merger.close()
        
        # 5. Limpeza (Opcional, mas recomendado)
        # Se quiser remover os PDFs de lote (intermediários) após a união:
        # for pdf_file in arquivos_pdf_gerados:
        #     os.remove(pdf_file)
        #     print(f"  -> Removido arquivo intermediário: {os.path.basename(pdf_file)}")
            
        print(f"\nSUCESSO: Catálogo final salvo como '{caminho_final_pdf}'")
    
    except Exception as e:
        print(f"ERRO CRÍTICO no processo de MERGE. Verifique se 'PyPDF2' está instalado corretamente.")
        print(f"Detalhe do Erro: {e}")

if __name__ == '__main__':
    unir_arquivos_pdf()