
# 📦 OFERTEIRO - Automação de Tabloide de Ofertas

## ✨ Visão Geral

O **OFERTEIRO** é um sistema de automação em Python projetado para gerar rapidamente um **Tabloide de Ofertas (ou Catálogo)** em formato PDF (`Tabloide_Ofertas.pdf`).

Ele processa dados de produtos a partir de uma planilha Excel (`dados.xlsx`) e os formata em um arquivo Word (`template.docx`), garantindo uma paginação consistente de **16 produtos por página** (tabela 4x4) e realizando a conversão final para PDF.

-----

## 🛠️ Pré-requisitos Essenciais

Para que a automação funcione corretamente, você deve ter os seguintes itens instalados no seu sistema:

1.  **Python 3.x:** (Versão 3.6 ou superior recomendada)
2.  **Software de Conversão PDF:** O script depende de um software externo para converter o Word para PDF.
      * **No Windows:** É obrigatório ter o **Microsoft Word** instalado.
      * **No Linux/macOS:** É obrigatório ter o **LibreOffice** instalado.

-----

## 📂 Estrutura do Projeto

Todos os arquivos listados abaixo devem estar localizados no **mesmo diretório**.

| Arquivo | Tipo | Função |
| :--- | :--- | :--- |
| `autodoc.py` | Python | Contém a **lógica de automação principal** (leitura dos dados, manipulação do Word e conversão). |
| `instalar_dependencias.bat` | Windows Batch | Script para **configurar o ambiente** Python e instalar as bibliotecas necessárias. |
| `dados.xlsx` | Planilha | **Fonte de dados** dos produtos (Nome, URL da Imagem, Preço). |
| `template.docx` | Word DOCX | O **modelo** do tabloide, contendo a tabela base (4x4) que será clonada para cada página. |
| `venv/` | Pasta | Ambiente virtual criado pelo script BAT, garantindo que as dependências fiquem isoladas. |

-----

## 1\. ⚙️ Configuração e Instalação

O processo de configuração é simplificado pelo script `instalar_dependencias.bat`.

### Passo 1: Executar o Instalador

1.  Dê um **duplo clique** no arquivo `instalar_dependencias.bat`.
2.  O script fará automaticamente:
      * Criação da pasta de Ambiente Virtual (`venv/`).
      * Ativação do ambiente.
      * Instalação de todas as dependências Python (`pandas`, `python-docx`, `requests`, `docx2pdf`, etc.).
3.  Aguarde até que a mensagem **"INSTALACAO CONCLUIDA COM SUCESSO\!"** apareça na tela. O terminal permanecerá ativo e pronto para a execução.

### Passo 2: Preparação dos Arquivos

#### A. Preparação da Planilha (`dados.xlsx`)

O script espera uma planilha sem cabeçalho e com as seguintes colunas obrigatórias:

| Coluna | Nome da Coluna | Conteúdo | Exemplo de Dado |
| :---: | :---: | :--- | :--- |
| **1** | `name` | Nome do Produto | Monitor Gamer 24" |
| **2** | `img_url` | URL da Imagem (web) | `http://link.com/img1.jpg` |
| **3** | `price` | Preço do Produto | `R$ 1.250,90` |

#### B. Preparação do Template (`template.docx`)

O template deve conter **apenas uma tabela** que será usada como modelo para todas as páginas.

  * **Tamanho Mínimo:** A tabela deve ter no mínimo **4 linhas x 4 colunas** para garantir a estrutura correta de 16 produtos por página.
  * **Formato:** O script irá clonar essa tabela, limpá-la e preenchê-la com os dados da planilha.

-----

## 2\. 🚀 Execução da Automação

Com as dependências instaladas e os arquivos de dados/template prontos, o processo de geração é simples.

### 4.1. 🏃 Etapa 1: Gerar Tabloide (DOCX Único)

A partir do terminal onde o `instalar_dependencias.bat` foi executado:

```bash
python autodoc.py
```

**Saída Esperada:**

  * O arquivo `Documentos_Gerados/Tabloide_Ofertas.docx` será criado com todos os produtos.
  * O script tentará, em seguida, iniciar a conversão automática para PDF (Etapa 2).

### 4.2. 💾 Etapa 2: Conversão Automática para PDF

O script utiliza a biblioteca `docx2pdf`, que, por sua vez, usa o **MS Word** (Windows) ou **LibreOffice** (Linux/macOS) instalado para realizar a conversão.

  * **Resultado:** O arquivo final `Documentos_Gerados/Tabloide_Ofertas.pdf` será gerado.

-----

## ❓ Solução de Problemas Comuns

| Problema | Causa Mais Comum | Solução |
| :--- | :--- | :--- |
| **Falha na Conversão para PDF** | Falta do MS Word/LibreOffice ou problema de permissão. | 1. Certifique-se de que o MS Word (Win) ou LibreOffice (Lin/Mac) está instalado. 2. Se a automação falhar, faça a **Conversão Manual** (veja abaixo). |
| **Falha ao salvar/permissão negada** | O arquivo DOCX está aberto ou em uso. | Feche o arquivo `Documentos_Gerados/Tabloide_Ofertas.docx` e execute `python autodoc.py` novamente. |
| **`pip` não é reconhecido** | Python/Pip não está no PATH global ou o ambiente virtual não está ativo. | Execute o script `instalar_dependencias.bat` novamente para garantir que o ambiente seja ativado. |

### ⚠️ Conversão Manual (Alternativa)

Se a automação falhar na conversão para PDF, siga estes passos:

1.  Abra o arquivo gerado: `Documentos_Gerados/Tabloide_Ofertas.docx`.
2.  Use a função **"Salvar Como"** (ou "Exportar") do seu editor de texto.
3.  Selecione o formato **PDF** e salve-o na mesma pasta.

-----
