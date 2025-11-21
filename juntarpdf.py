import os
import glob
from PyPDF2 import PdfMerger # Necessário: pip install pypdf2

# --- Configurações ---
SAIDA_PASTA = 'Documentos_Gerados'
NOME_FINAL_PDF = 'Catalogo_Final_UNIDO.pdf'

def obter_numero_lote(caminho_arquivo):
    """Extrai o número do lote do nome do arquivo para ordenação numérica."""
    # Ex: 'Documentos_Gerados/Catalogo_Lote_10.pdf' -> 10
    nome_base = os.path.basename(caminho_arquivo)
    
    # Remove a extensão '.pdf' e divide
    nome_sem_extensao = os.path.splitext(nome_base)[0]
    
    try:
        numero_str = nome_sem_extensao.split('_')[-1]
        return int(numero_str) # Converte para INTEIRO para ordenação numérica correta
    except:
        # Retorna 0 (ou um valor seguro) em caso de erro no nome
        return 0

def unir_arquivos_pdf():
    """
    Escaneia a pasta de saída, encontra todos os PDFs gerados e os une 
    em um único arquivo final.
    """
    
    # 1. Verificar se a pasta existe
    if not os.path.exists(SAIDA_PASTA):
        print(f"ERRO: A pasta de saída '{SAIDA_PASTA}' não foi encontrada.")
        print("Execute o script de automação de documentos primeiro.")
        return # Garante que o bloco 'if' tem instruções identadas

    # 2. Encontrar todos os PDFs de Lote
    padrao_busca = os.path.join(SAIDA_PASTA, 'Catalogo_Lote_*.pdf')
    
    # --- CORREÇÃO DA ORDENAÇÃO NUMÉRICA ---
    # Usa a função 'obter_numero_lote' como chave de ordenação
    arquivos_pdf_gerados = sorted(glob.glob(padrao_busca), key=obter_numero_lote) 
    
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
        # for pdf_file in arquivos_pdf_gerados:
        #    os.remove(pdf_file)
        #    print(f"  -> Removido arquivo intermediário: {os.path.basename(pdf_file)}")
            
        print(f"\nSUCESSO: Catálogo final salvo como '{caminho_final_pdf}'")
        
    except Exception as e:
        print(f"ERRO CRÍTICO no processo de MERGE. Verifique se 'PyPDF2' está instalado corretamente.")
        print(f"Detalhe do Erro: {e}")

if __name__ == '__main__':
    unir_arquivos_pdf()
