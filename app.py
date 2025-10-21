import streamlit as st
import pandas as pd
from tabula import read_pdf
from io import BytesIO
import tempfile
import os
import subprocess
import zipfile
import PyPDF2
import re

st.set_page_config(
    page_title="Conversor PDF para Excel",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Conversor de PDF para Excel")
st.markdown("Extraia tabelas de arquivos PDF e converta para Excel facilmente!")

# Verificar se Java está instalado
try:
    result = subprocess.run(['java', '-version'], capture_output=True, text=True)
    java_ok = True
except FileNotFoundError:
    java_ok = False
    st.error("❌ Java não encontrado! O tabula-py requer Java instalado.")
    st.info("📥 Baixe o Java em: https://www.java.com/download/")
    st.stop()

# Inicializar session state
if 'limpar' not in st.session_state:
    st.session_state.limpar = False

# Upload dos arquivos
col1, col2 = st.columns([3, 1])
with col1:
    uploaded_files = st.file_uploader(
        "Escolha um ou mais arquivos PDF",
        type=['pdf'],
        accept_multiple_files=True,
        help="Selecione os arquivos PDF que contêm as tabelas",
        key='uploader' if not st.session_state.limpar else 'uploader_limpo'
    )

with col2:
    st.write("")  # Espaçamento
    if st.button("🗑️ Limpar arquivos", use_container_width=True, type="secondary"):
        st.session_state.limpar = not st.session_state.limpar
        st.rerun()

if uploaded_files:
    # Opções de configuração
    st.sidebar.header("⚙️ Configurações")
    
    metodo_extracao = st.sidebar.selectbox(
        "Método de extração",
        ["Automático", "Lattice (tabelas com bordas)", "Stream (tabelas sem bordas)", "Texto (extração inteligente)"],
        help="Lattice: melhor para tabelas com linhas visíveis. Stream: melhor para tabelas alinhadas por espaços. Texto: extrai e estrutura texto do PDF."
    )
    
    ajustar_headers = st.sidebar.checkbox(
        "Ajustar cabeçalhos automaticamente",
        value=False,
        help="Combina a primeira linha com o cabeçalho quando necessário"
    )
    mostrar_preview = st.sidebar.checkbox(
        "Mostrar prévia das tabelas",
        value=True
    )
    
    # Função para ajustar header
    def ajusta_header(df):
        novo_header = []
        for header_label, first_row_label in zip(df.columns, df.iloc[0]):
            if header_label.startswith('Unnamed'):
                novo_header.append(str(first_row_label))
            else:
                novo_header.append(f'{header_label} {first_row_label}')
        df.columns = novo_header
        return df.iloc[1:].reset_index(drop=True)
    
    # Função para extrair texto e estruturar
    def extrair_texto_estruturado(pdf_path):
        """Extrai texto do PDF e tenta estruturar em tabela"""
        try:
            with open(pdf_path, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                texto_completo = ""
                for page in reader.pages:
                    texto_completo += page.extract_text() + "\n"
            
            linhas = texto_completo.split('\n')
            
            # Dicionário para armazenar dados estruturados
            dados_pedido = {
                'Informações Gerais': [],
                'Produtos': []
            }
            
            # --- EXTRAÇÃO DE INFORMAÇÕES GERAIS ---
            info_geral = {}
            for linha in linhas:
                linha_limpa = linha.strip()
                
                # Número do Pedido
                if 'Número do Pedido' in linha_limpa or 'Pedido:' in linha_limpa:
                    match = re.search(r'(?:Número do Pedido:|Pedido:)\s*(\d+)', linha_limpa)
                    if match:
                        info_geral['Número do Pedido'] = match.group(1)
                
                # Fornecedor
                if 'Fornecedor:' in linha_limpa:
                    match = re.search(r'Fornecedor:\s*\d+\s*(.+?)(?:,\s*CNPJ|$)', linha_limpa)
                    if match:
                        info_geral['Fornecedor'] = match.group(1).strip()
                
                # CNPJ Fornecedor
                if 'CNPJ:' in linha_limpa and 'Fornecedor' in linha_limpa:
                    match = re.search(r'CNPJ:\s*([\d\.\/\-]+)', linha_limpa)
                    if match:
                        info_geral['CNPJ Fornecedor'] = match.group(1)
                
                # Empresa
                if 'Empresa:' in linha_limpa and 'CNPJ' not in linha_limpa:
                    match = re.search(r'Empresa:\s*\d+\s*(.+?)(?:,|$)', linha_limpa)
                    if match:
                        info_geral['Empresa'] = match.group(1).strip()
                
                # Datas
                if 'Dt. Pedido' in linha_limpa:
                    match = re.search(r'Dt\.\s*Pedido:\s*([\d\/]+)', linha_limpa)
                    if match:
                        info_geral['Data Pedido'] = match.group(1)
                
                if 'Dt. Entrega' in linha_limpa:
                    match = re.search(r'Dt\.\s*Entrega:\s*([\d\/]+)', linha_limpa)
                    if match:
                        info_geral['Data Entrega'] = match.group(1)
                
                # Forma de Pagamento
                if 'Forma Pgto' in linha_limpa or 'Forma de Pagamento' in linha_limpa:
                    match = re.search(r'Forma\s+Pgto:\s*(.+?)(?:,\s*Espécie|$)', linha_limpa)
                    if match:
                        info_geral['Forma Pagamento'] = match.group(1).strip()
                
                # Frete
                if 'Frete:' in linha_limpa and 'Forma' not in linha_limpa:
                    match = re.search(r'Frete:\s*(\w+)', linha_limpa)
                    if match:
                        info_geral['Frete'] = match.group(1)
            
            # Adicionar informações gerais ao resultado
            if info_geral:
                for chave, valor in info_geral.items():
                    dados_pedido['Informações Gerais'].append({
                        'Campo': chave,
                        'Valor': valor
                    })
            
            # --- EXTRAÇÃO DE PRODUTOS ---
            # Procurar linha de cabeçalho de produtos
            idx_header = -1
            for i, linha in enumerate(linhas):
                if 'Código' in linha and 'Descrição' in linha and ('Qtde' in linha or 'Quantidade' in linha):
                    idx_header = i
                    break
            
            # Se encontrou cabeçalho, processar produtos
            if idx_header != -1:
                # Extrair produtos (próximas linhas após o cabeçalho)
                for i in range(idx_header + 1, len(linhas)):
                    linha = linhas[i].strip()
                    
                    # Parar em linhas de rodapé
                    if any(x in linha.lower() for x in ['recebimento', 'comprador', 'vendedor', 'obrigatório', '---', 'pg:']):
                        break
                    
                    if not linha or len(linha) < 10:
                        continue
                    
                    # Tentar extrair dados do produto
                    partes = re.split(r'\s{2,}', linha)
                    
                    if len(partes) >= 3:
                        produto = {}
                        
                        # Código do produto (geralmente números no início)
                        if partes[0].isdigit():
                            produto['Código'] = partes[0]
                        
                        # Código de barras (geralmente 13 dígitos)
                        for parte in partes:
                            if parte.isdigit() and len(parte) >= 12:
                                produto['Código Barras'] = parte
                                break
                        
                        # Descrição (texto mais longo)
                        descricoes = [p for p in partes if not p.replace('.', '').replace(',', '').isdigit() and len(p) > 3]
                        if descricoes:
                            produto['Descrição'] = ' '.join(descricoes[:2])
                            if len(descricoes) > 2:
                                produto['Marca'] = descricoes[2]
                        
                        # Quantidade (número com vírgula)
                        for parte in partes:
                            if ',' in parte or '.' in parte:
                                try:
                                    # Verifica se parece com quantidade (ex: 20,000 ou 20.000)
                                    num_str = parte.replace('.', '').replace(',', '.')
                                    num = float(num_str)
                                    if num < 10000:  # Quantidade geralmente não é muito grande
                                        produto['Quantidade'] = parte
                                        break
                                except:
                                    pass
                        
                        # Valores monetários (procurar por números com vírgula e pelo menos 2 casas decimais)
                        valores = []
                        for p in partes:
                            if ',' in p or '.' in p:
                                try:
                                    # Verificar se tem pelo menos 2 dígitos decimais
                                    if ',' in p:
                                        partes_decimal = p.split(',')
                                    else:
                                        partes_decimal = p.split('.')
                                    
                                    if len(partes_decimal) == 2 and len(partes_decimal[1]) >= 2:
                                        valores.append(p)
                                except:
                                    pass
                        
                        if len(valores) >= 2:
                            produto['Preço Unitário'] = valores[0]
                            produto['Valor Total'] = valores[-1]
                        
                        # Embalagem (CX/20, UN, etc)
                        for parte in partes:
                            if '/' in parte or parte.upper() in ['CX', 'UN', 'PC', 'KG', 'LT']:
                                produto['Embalagem'] = parte
                                break
                        
                        if len(produto) >= 2:
                            dados_pedido['Produtos'].append(produto)
            
            # Criar DataFrames
            dfs = []
            
            if dados_pedido['Informações Gerais']:
                df_info = pd.DataFrame(dados_pedido['Informações Gerais'])
                dfs.append(df_info)
            
            if dados_pedido['Produtos']:
                df_produtos = pd.DataFrame(dados_pedido['Produtos'])
                dfs.append(df_produtos)
            
            # Se não conseguiu estruturar, retornar texto bruto organizado
            if not dfs:
                dados_genericos = []
                for linha in linhas:
                    linha = linha.strip()
                    if not linha or len(linha) < 5:
                        continue
                    
                    campos = re.split(r'\s{2,}', linha)
                    if len(campos) >= 3:
                        dados_genericos.append(campos)
                
                if dados_genericos and len(dados_genericos) > 1:
                    max_cols = max(len(row) for row in dados_genericos)
                    
                    dados_padronizados = []
                    for row in dados_genericos:
                        while len(row) < max_cols:
                            row.append('')
                        dados_padronizados.append(row[:max_cols])
                    
                    df = pd.DataFrame(dados_padronizados[1:], columns=dados_padronizados[0])
                    dfs.append(df)
                else:
                    df = pd.DataFrame({'Conteúdo': [l.strip() for l in linhas if l.strip()]})
                    dfs.append(df)
            
            return dfs
            
        except Exception as e:
            st.error(f"Erro ao extrair texto: {str(e)}")
            return []
    
    # Processar cada arquivo
    resultados = []
    
    for uploaded_file in uploaded_files:
        st.divider()
        st.subheader(f"📄 {uploaded_file.name}")
        
        # Salvar arquivo temporariamente
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
            tmp_file.write(uploaded_file.read())
            tmp_path = tmp_file.name
        
        try:
            with st.spinner(f'Extraindo tabelas de {uploaded_file.name}...'):
                # Aplicar método selecionado
                if metodo_extracao == "Texto (extração inteligente)":
                    tabelas = extrair_texto_estruturado(tmp_path)
                elif metodo_extracao == "Lattice (tabelas com bordas)":
                    tabelas = read_pdf(tmp_path, pages='all', lattice=True)
                elif metodo_extracao == "Stream (tabelas sem bordas)":
                    tabelas = read_pdf(tmp_path, pages='all', stream=True)
                else:  # Automático
                    # Tentar extrair tabelas com diferentes métodos
                    try:
                        tabelas = read_pdf(tmp_path, pages='all', lattice=True)
                        if not tabelas or all(df.empty for df in tabelas):
                            raise Exception("Nenhuma tabela encontrada com método lattice")
                    except:
                        try:
                            tabelas = read_pdf(tmp_path, pages='all', stream=True)
                            if not tabelas or all(df.empty for df in tabelas):
                                raise Exception("Nenhuma tabela encontrada com método stream")
                        except:
                            # Tentar extração de texto como último recurso
                            tabelas = extrair_texto_estruturado(tmp_path)
                            if not tabelas:
                                tabelas = read_pdf(tmp_path, pages='all', guess=True, multiple_tables=True)
            
            if not tabelas or len(tabelas) == 0:
                st.warning(f"⚠️ Nenhuma tabela foi encontrada em {uploaded_file.name}")
                st.info("💡 Dica: Este PDF pode conter texto não estruturado ou tabelas em formato de imagem. Tente converter o PDF ou usar OCR.")
                continue
            
            st.success(f'✅ Foram encontradas {len(tabelas)} tabelas!')
            
            # Criar Excel em memória
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                tabelas_validas = 0
                
                for i, df in enumerate(tabelas, 1):
                    # Pular tabelas vazias ou muito pequenas
                    if df.empty or len(df.columns) < 2:
                        continue
                    
                    # Ajustar headers se solicitado
                    if ajustar_headers:
                        for col_name in df.columns:
                            if 'Unnamed' in str(col_name):
                                df = ajusta_header(df)
                                break
                    
                    # Escrever no Excel
                    sheet_name = f'Tabela {i}'
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    tabelas_validas += 1
                    
                    # Mostrar preview
                    if mostrar_preview:
                        with st.expander(f"📋 Tabela {i} - {len(df)} linhas x {len(df.columns)} colunas"):
                            st.dataframe(df, use_container_width=True)
            
            output.seek(0)
            
            # Armazenar resultado
            resultados.append({
                'nome': uploaded_file.name.replace('.pdf', ''),
                'dados': output.getvalue(),
                'tabelas': tabelas_validas
            })
            
            # Botão de download individual
            col_a, col_b, col_c = st.columns([1, 2, 1])
            with col_b:
                st.download_button(
                    label=f"⬇️ Baixar {uploaded_file.name.replace('.pdf', '')}.xlsx",
                    data=output.getvalue(),
                    file_name=f"{uploaded_file.name.replace('.pdf', '')}_tabelas.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key=f"download_{uploaded_file.name}"
                )
            
            st.info(f"📊 Total de tabelas válidas exportadas: {tabelas_validas}")
            
        except Exception as e:
            st.error(f"❌ Erro ao processar {uploaded_file.name}: {str(e)}")
            st.info("💡 Dica: Certifique-se de que o PDF contém tabelas extraíveis (não imagens de tabelas)")
        
        finally:
            # Limpar arquivo temporário
            if os.path.exists(tmp_path):
                os.unlink(tmp_path)
    
    # Botão para baixar todos os arquivos em ZIP (se houver mais de um)
    if len(resultados) > 1:
        st.divider()
        st.subheader("📦 Download em Lote")
        
        # Criar ZIP com todos os arquivos
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for resultado in resultados:
                zip_file.writestr(
                    f"{resultado['nome']}_tabelas.xlsx",
                    resultado['dados']
                )
        
        zip_buffer.seek(0)
        
        col_x, col_y, col_z = st.columns([1, 2, 1])
        with col_y:
            st.download_button(
                label=f"📦 Baixar todos ({len(resultados)} arquivos) em ZIP",
                data=zip_buffer.getvalue(),
                file_name="tabelas_convertidas.zip",
                mime="application/zip",
                use_container_width=True
            )

else:
    # Instruções quando nenhum arquivo foi carregado
    st.info("👆 Faça upload de um ou mais arquivos PDF para começar")
    
    with st.expander("ℹ️ Como usar"):
        st.markdown("""
        1. **Faça upload** de um ou mais arquivos PDF contendo tabelas
        2. Aguarde a **extração automática** das tabelas de cada arquivo
        3. **Configure** as opções na barra lateral (opcional)
        4. **Visualize** as tabelas extraídas de cada PDF
        5. **Baixe** os arquivos Excel individualmente ou todos em ZIP
        6. Use o botão **"Limpar arquivos"** para começar novamente
        
        **Requisitos:**
        - O PDF deve conter tabelas estruturadas
        - Tabelas em formato de imagem não serão extraídas
        """)
    
    with st.expander("📦 Dependências necessárias"):
        st.code("""
pip install streamlit pandas tabula-py openpyxl PyPDF2
        """, language="bash")
        st.warning("⚠️ **Importante**: Instale `tabula-py` (não `tabula`) e certifique-se de ter Java instalado!")
        st.markdown("Verificar Java: `java -version`")