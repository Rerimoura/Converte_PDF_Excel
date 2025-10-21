import streamlit as st
import pandas as pd
from tabula import read_pdf
from io import BytesIO
import tempfile
import os
import subprocess
import zipfile

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
if 'arquivos_processados' not in st.session_state:
    st.session_state.arquivos_processados = []

# Upload dos arquivos
col1, col2 = st.columns([3, 1])
with col1:
    uploaded_files = st.file_uploader(
        "Escolha um ou mais arquivos PDF",
        type=['pdf'],
        accept_multiple_files=True,
        help="Selecione os arquivos PDF que contêm as tabelas"
    )

with col2:
    st.write("")  # Espaçamento
    if st.button("🗑️ Limpar arquivos", use_container_width=True, type="secondary"):
        st.session_state.arquivos_processados = []
        st.rerun()

if uploaded_files:
    # Opções de configuração
    st.sidebar.header("⚙️ Configurações")
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
                # Extrair tabelas
                tabelas = read_pdf(tmp_path, pages='all')
            
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
pip install streamlit pandas tabula-py openpyxl
        """, language="bash")
        st.warning("⚠️ **Importante**: Instale `tabula-py` (não `tabula`) e certifique-se de ter Java instalado!")
        st.markdown("Verificar Java: `java -version`")