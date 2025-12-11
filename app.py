import streamlit as st
import pandas as pd
from io import BytesIO
import tempfile
import os
import zipfile
import re
from extractors import PdfPlumberExtractor, TabulaExtractor, TextExtractor, RedeBizExtractor, MondelezExtractor, SilveiraExtractor, TABULA_AVAILABLE

st.set_page_config(
    page_title="Conversor PDF para Excel",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Conversor de PDF para Excel")
st.markdown("Extraia tabelas de arquivos PDF e converta para Excel facilmente!")

# Número do WhatsApp (formato: código do país + DDD + número)
WHATSAPP_NUMBER = "553492182544"

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
    
    opcoes_extracao = [
        "REDE BIZ (Pedidos TOTVS)",
        "Rede Biz - KAMEL",
        "Silveira Supermercado",
        "PDFPlumber (recomendado)", 
        "Texto (extração inteligente)"
    ]
    if TABULA_AVAILABLE:
        opcoes_extracao.extend(["Lattice (Tabula - bordas)", "Stream (Tabula - sem bordas)", "Automático (Tabula)"])
    
    metodo_extracao = st.sidebar.selectbox(
        "Método de extração",
        opcoes_extracao,
        help="Escolha o algoritmo de extração."
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
            if str(header_label).startswith('Unnamed'):
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
                tabelas = []
                
                # Selecionar extrator
                extractor = None
                
                if "REDE BIZ" in metodo_extracao:
                    extractor = RedeBizExtractor()
                elif "KAMEL" in metodo_extracao:
                    extractor = MondelezExtractor()
                elif "Silveira" in metodo_extracao:
                    extractor = SilveiraExtractor()
                elif "PDFPlumber" in metodo_extracao:
                    extractor = PdfPlumberExtractor()
                elif "Texto" in metodo_extracao:
                    extractor = TextExtractor()
                elif "Lattice" in metodo_extracao:
                    extractor = TabulaExtractor(mode="lattice")
                elif "Stream" in metodo_extracao:
                    extractor = TabulaExtractor(mode="stream")
                elif "Automático" in metodo_extracao:
                    extractor = TabulaExtractor(mode="auto")
                
                if extractor:
                    tabelas = extractor.extract(tmp_path)
            
            if not tabelas or len(tabelas) == 0:
                st.warning(f"⚠️ Nenhuma tabela foi encontrada em {uploaded_file.name}")
                if hasattr(extractor, 'debug_text') and extractor.debug_text:
                    with st.expander("🔍 Debug: Ver texto extraído do PDF"):
                        st.text_area("Conteúdo bruto:", extractor.debug_text, height=300)
                        st.info("Copie este texto e envie para o desenvolvedor adaptar o extrator.")
                st.info("💡 Dica: Tente outro método de extração.")
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
                        try:
                            df = ajusta_header(df)
                        except:
                            pass
                    
                    # Converter EAN para número se existir
                    if 'EAN' in df.columns:
                        # Forçar conversão para números, erros viram NaN (que o Excel trata como vazio)
                        # Mas se quiser manter zeros à esquerda, teria que ser string.
                        # O usuário PEDIU para ser número.
                        df['EAN'] = pd.to_numeric(df['EAN'], errors='coerce')

                    # Escrever no Excel
                    sheet_name = f'Tabela {i}'
                    # Garantir que sheet_name não exceda 31 chars
                    sheet_name = sheet_name[:31]
                    
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    tabelas_validas += 1
                    
                    # Mostrar preview
                    if mostrar_preview:
                        with st.expander(f"📋 Tabela {i} - {len(df)} linhas x {len(df.columns)} colunas"):
                            st.dataframe(df, use_container_width=True)
            
            output.seek(0)
            
            if tabelas_validas > 0:
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
            else:
                st.warning("⚠️ Tabelas encontradas, mas vazias ou inválidas.")
            
        except Exception as e:
            st.error(f"❌ Erro ao processar {uploaded_file.name}: {str(e)}")
        
        finally:
            # Limpar arquivo temporário
            if os.path.exists(tmp_path):
                try:
                    os.unlink(tmp_path)
                except:
                    pass
    
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

    # Botão de WhatsApp para suporte geral
    col_sup1, col_sup2, col_sup3 = st.columns([1, 2, 1])
    with col_sup2:
        st.markdown("### 💬 Precisa de ajuda?")
        mensagem_suporte = "Olá! Preciso de ajuda com o conversor de PDF para Excel."
        mensagem_encoded = mensagem_suporte.replace(" ", "%20")
        whatsapp_url = f"https://wa.me/{WHATSAPP_NUMBER}?text={mensagem_encoded}"
        
        st.markdown(
            f"""
            <a href="{whatsapp_url}" target="_blank">
                <button style="
                    background-color: #25D366;
                    color: white;
                    padding: 12px 24px;
                    border: none;
                    border-radius: 8px;
                    cursor: pointer;
                    font-size: 16px;
                    font-weight: bold;
                    width: 100%;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    gap: 8px;
                ">
                    💬 Mande mensagem para {WHATSAPP_NUMBER}
                </button>
            </a>
            """,
            unsafe_allow_html=True
        )
    
    st.divider()
    
    with st.expander("ℹ️ Como usar"):
        st.markdown("""
        1. **Faça upload** de um ou mais arquivos PDF contendo tabelas
        2. **Escolha o método** de extração (PDFPlumber é o padrão e não requer Java)
        3. **Visualize** as tabelas extraídas
        4. **Baixe** os arquivos Excel
        
        **Dicas:**
        - **PDFPlumber**: Ótimo para a maioria dos casos, rápido e não precisa de Java.
        - **Texto**: Use para notas fiscais ou pedidos com layout específico, ou se outros métodos falharem.
        - **Tabula (Lattice)**: Se tiver Java instalado, use para tabelas com bordas bem definidas.
        """)