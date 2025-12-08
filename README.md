# Conversor PDF para Excel

Este é um aplicativo Streamlit que extrai tabelas de arquivos PDF e as converte para o formato Excel (XLSX).

## Funcionalidades

- Upload de múltiplos arquivos PDF
- Extração automática de tabelas
- Suporte a múltiplos métodos de extração:
    - **PDFPlumber** (Padrão, não requer Java)
    - **REDE BIZ** (Extração específica para pedidos TOTVS)
    - **Tabula** (Lattice/Stream - requer Java)
    - **Texto** (Extração baseada em Regex)
- Preview das tabelas extraídas
- Download individual ou em lote (ZIP)

## Como executar localmente

1. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```

2. Execute o aplicativo:
   ```bash
   streamlit run app.py
   ```

## Deploy no Streamlit Cloud

Para publicar este aplicativo no Streamlit Cloud:

1. Crie um repositório no GitHub.
2. Faça upload dos seguintes arquivos:
   - `app.py`
   - `extractors.py`
   - `requirements.txt`
   - `packages.txt`
3. Conecte sua conta do Streamlit Cloud ao GitHub.
4. Selecione o repositório e o arquivo `app.py`.
5. Em "Advanced settings", certifique-se que a versão do Python é compatível (ex: 3.9 ou superior).

**Nota:** O arquivo `packages.txt` é necessário para instalar o Java (default-jre) no ambiente do Streamlit Cloud, que é requerido para o método de extração Tabula.
