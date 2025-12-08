import pandas as pd
import pdfplumber
import PyPDF2
import re
from io import BytesIO
try:
    from tabula import read_pdf
    TABULA_AVAILABLE = True
except ImportError:
    TABULA_AVAILABLE = False
import subprocess

class PdfExtractor:
    """Base class for PDF extractors"""
    def extract(self, file_path):
        raise NotImplementedError("Subclasses must implement extract method")

class PdfPlumberExtractor(PdfExtractor):
    """Extract tables using pdfplumber"""
    def extract(self, file_path):
        tabelas = []
        try:
            with pdfplumber.open(file_path) as pdf:
                for i, page in enumerate(pdf.pages, start=1):
                    tables = page.extract_tables()
                    for j, table in enumerate(tables, start=1):
                        if table:
                            # Clean table data
                            cleaned_table = []
                            for row in table:
                                # Filter out None values and replace with empty string
                                cleaned_row = [cell if cell is not None else "" for cell in row]
                                # Only add rows that have at least one non-empty cell
                                if any(str(c).strip() for c in cleaned_row):
                                    cleaned_table.append(cleaned_row)
                            
                            if len(cleaned_table) > 1:
                                df = pd.DataFrame(cleaned_table[1:], columns=cleaned_table[0])
                                tabelas.append(df)
        except Exception as e:
            print(f"Error in PdfPlumberExtractor: {e}")
        return tabelas

class TabulaExtractor(PdfExtractor):
    """Extract tables using tabula-py"""
    def __init__(self, mode="lattice"):
        self.mode = mode # lattice, stream, or auto

    def check_java(self):
        try:
            subprocess.run(['java', '-version'], capture_output=True, text=True)
            return True
        except FileNotFoundError:
            return False

    def extract(self, file_path):
        if not self.check_java():
            raise EnvironmentError("Java not found. Tabula requires Java.")
            
        if not TABULA_AVAILABLE:
            raise ImportError("tabula-py not installed.")

        tabelas = []
        try:
            if self.mode == "lattice":
                tabelas = read_pdf(file_path, pages='all', lattice=True)
            elif self.mode == "stream":
                tabelas = read_pdf(file_path, pages='all', stream=True)
            else: # auto
                try:
                    tabelas = read_pdf(file_path, pages='all', lattice=True)
                    if not tabelas or all(df.empty for df in tabelas):
                        raise Exception("Empty lattice result")
                except:
                    tabelas = read_pdf(file_path, pages='all', stream=True)
        except Exception as e:
            print(f"Error in TabulaExtractor: {e}")
        
        # Filter empty dataframes
        if tabelas:
            tabelas = [df for df in tabelas if not df.empty]
            
        return tabelas

class TextExtractor(PdfExtractor):
    """Extract structured data from text using regex"""
    def extract(self, file_path):
        try:
            with open(file_path, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                texto_completo = ""
                for page in reader.pages:
                    text = page.extract_text()
                    if text:
                        texto_completo += text + "\n"
            
            linhas = texto_completo.split('\n')
            return self._process_text(linhas)
        except Exception as e:
            print(f"Error in TextExtractor: {e}")
            return []

    def _process_text(self, linhas):
        # ... logic from extrair_texto_estruturado ...
        # Copied from original app.py and adapted
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
        
        if info_geral:
            for chave, valor in info_geral.items():
                dados_pedido['Informações Gerais'].append({
                    'Campo': chave,
                    'Valor': valor
                })
        
        # --- EXTRAÇÃO DE PRODUTOS ---
        idx_header = -1
        for i, linha in enumerate(linhas):
            if 'Código' in linha and 'Descrição' in linha and ('Qtde' in linha or 'Quantidade' in linha):
                idx_header = i
                break
        
        if idx_header != -1:
            for i in range(idx_header + 1, len(linhas)):
                linha = linhas[i].strip()
                if any(x in linha.lower() for x in ['recebimento', 'comprador', 'vendedor', 'obrigatório', '---', 'pg:']):
                    break
                if not linha or len(linha) < 10:
                    continue
                
                partes = re.split(r'\s{2,}', linha)
                if len(partes) >= 3:
                    produto = {}
                    if partes[0].isdigit():
                        produto['Código'] = partes[0]
                    for parte in partes:
                        if parte.isdigit() and len(parte) >= 12:
                            produto['Código Barras'] = parte
                            break
                    descricoes = [p for p in partes if not p.replace('.', '').replace(',', '').isdigit() and len(p) > 3]
                    if descricoes:
                        produto['Descrição'] = ' '.join(descricoes[:2])
                        if len(descricoes) > 2:
                            produto['Marca'] = descricoes[2]
                    for parte in partes:
                        if ',' in parte or '.' in parte:
                            try:
                                num_str = parte.replace('.', '').replace(',', '.')
                                num = float(num_str)
                                if num < 10000:
                                    produto['Quantidade'] = parte
                                    break
                            except:
                                pass
                    valores = []
                    for p in partes:
                        if ',' in p or '.' in p:
                            try:
                                if ',' in p: partes_decimal = p.split(',')
                                else: partes_decimal = p.split('.')
                                if len(partes_decimal) == 2 and len(partes_decimal[1]) >= 2:
                                    valores.append(p)
                            except: pass
                    if len(valores) >= 2:
                        produto['Preço Unitário'] = valores[0]
                        produto['Valor Total'] = valores[-1]
                    for parte in partes:
                        if '/' in parte or parte.upper() in ['CX', 'UN', 'PC', 'KG', 'LT']:
                            produto['Embalagem'] = parte
                            break
                    if len(produto) >= 2:
                        dados_pedido['Produtos'].append(produto)
        
        dfs = []
        if dados_pedido['Informações Gerais']:
            dfs.append(pd.DataFrame(dados_pedido['Informações Gerais']))
        if dados_pedido['Produtos']:
            dfs.append(pd.DataFrame(dados_pedido['Produtos']))
            
        if not dfs:
             # Fallback logic for unstructured text
            dados_genericos = []
            for linha in linhas:
                linha = linha.strip()
                if not linha or len(linha) < 5: continue
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
                
                # Check duplicates in headers
                headers = dados_padronizados[0]
                if len(headers) != len(set(headers)):
                     # Make headers unique
                    counts = {}
                    new_headers = []
                    for h in headers:
                        if h in counts:
                            counts[h] += 1
                            new_headers.append(f"{h}_{counts[h]}")
                        else:
                            counts[h] = 0
                            new_headers.append(h)
                    headers = new_headers

                df = pd.DataFrame(dados_padronizados[1:], columns=headers)
                dfs.append(df)
            else:
                dfs.append(pd.DataFrame({'Conteúdo': [l.strip() for l in linhas if l.strip()]}))
                
        return dfs

class RedeBizExtractor(PdfExtractor):
    """Extract structured data from REDE BIZ purchase orders (TOTVS format)"""
    
    def extract(self, file_path):
        try:
            # Extract all text from PDF
            with open(file_path, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                texto_completo = ""
                for page in reader.pages:
                    text = page.extract_text()
                    if text:
                        texto_completo += text + "\n"
            
            linhas = texto_completo.split('\n')
            return self._process_redebiz_text(linhas)
        except Exception as e:
            print(f"Error in RedeBizExtractor: {e}")
            return []
    
    def _process_redebiz_text(self, linhas):
        """Process REDE BIZ TOTVS purchase order format"""
        pedidos = []
        produtos_list = []
        
        current_pedido = None
        current_produto = None
        in_produtos_section = False
        
        for i, linha in enumerate(linhas):
            linha_limpa = linha.strip()
            
            # Detectar início de novo pedido
            if 'PEDIDO DE COMPRAS' in linha_limpa:
                # Extrair número do pedido
                match = re.search(r'PEDIDO DE COMPRAS\s+(\S+)', linha_limpa)
                numero_pedido = match.group(1) if match else ''
                
                # Verificar se é realmente um NOVO pedido (número diferente)
                if current_pedido and current_pedido['Número do Pedido'] != numero_pedido:
                    # Salvar pedido anterior apenas se o número for diferente
                    pedidos.append(current_pedido)
                    
                    # Criar novo pedido
                    current_pedido = {
                        'Número do Pedido': numero_pedido,
                        'Fornecedor': '',
                        'CNPJ Fornecedor': '',
                        'Cliente': '',
                        'CNPJ Cliente': '',
                        'Endereço Entrega': '',
                        'Cidade Entrega': '',
                        'Data Limite Entrega': '',
                        'Condição Frete': '',
                        'Data Emissão': '',
                        'Valor Total': ''
                    }
                    in_produtos_section = False
                elif not current_pedido:
                    # Primeiro pedido do documento
                    current_pedido = {
                        'Número do Pedido': numero_pedido,
                        'Fornecedor': '',
                        'CNPJ Fornecedor': '',
                        'Cliente': '',
                        'CNPJ Cliente': '',
                        'Endereço Entrega': '',
                        'Cidade Entrega': '',
                        'Data Limite Entrega': '',
                        'Condição Frete': '',
                        'Data Emissão': '',
                        'Valor Total': ''
                    }
                    in_produtos_section = False
                # Se for o mesmo pedido (mesma página ou continuação), apenas continua processando
                continue
            
            if not current_pedido:
                continue
            
            # Extrair informações do cabeçalho
            if 'R. Social' in linha_limpa and 'REDE BIZ' in linha_limpa:
                # Fornecedor
                match = re.search(r'REDE BIZ SERVICOS E DISTRIBUICAO DE PRO', linha_limpa)
                if match:
                    current_pedido['Fornecedor'] = 'REDE BIZ SERVICOS E DISTRIBUICAO'
                
                # Cliente
                match = re.search(r'R\. Social SUPERMERCADO JB[^\n]*?LTDA', linha_limpa)
                if match:
                    current_pedido['Cliente'] = match.group(0).replace('R. Social ', '')
            
            if 'CNPJ' in linha_limpa and 'REDE BIZ' not in linha_limpa:
                # CNPJ do Cliente (formato invertido no PDF: "CNPJ -27 18.510.982/0001")
                # Padrão: "CNPJ", espaço, sufixo (ex: -27), espaço, corpo (ex: 18...)
                match = re.search(r'CNPJ\s+([-\d]{2,4})\s+([\d\.\/]{10,18})', linha_limpa)
                if match:
                    # Formato encontrado: sufixo (grupo 1) e corpo (grupo 2)
                    # Montar na ordem correta: corpo + sufixo
                    current_pedido['CNPJ Cliente'] = match.group(2) + match.group(1)
                else:
                    # Tenta formato normal se o específico não der match
                    match = re.search(r'CNPJ\s+([\d\.\-\/]+)', linha_limpa)
                    if match:
                        current_pedido['CNPJ Cliente'] = match.group(1)
            
            if 'CNPJ' in linha_limpa and 'REDE BIZ' in linha_limpa:
                # CNPJ do Fornecedor
                match = re.search(r'CNPJ\s+([\d\.\-\/]+)', linha_limpa)
                if match:
                    current_pedido['CNPJ Fornecedor'] = match.group(1)
            
            if 'Data limite para entrega' in linha_limpa:
                match = re.search(r'Data limite para entrega\s+([\d\/]+)', linha_limpa)
                if match:
                    current_pedido['Data Limite Entrega'] = match.group(1)
            
            if 'Condi' in linha_limpa and 'o do frete' in linha_limpa:
                match = re.search(r'o do frete\s+(\w+)', linha_limpa)
                if match:
                    current_pedido['Condição Frete'] = match.group(1)
            
            if 'Data da emiss' in linha_limpa:
                match = re.search(r'Data da emiss.+?\s+([\d\/]+)', linha_limpa)
                if match:
                    current_pedido['Data Emissão'] = match.group(1)
            
            if 'Valor total do pedido' in linha_limpa:
                match = re.search(r'Valor total do pedido\s+([\d\.,]+)', linha_limpa)
                if match:
                    current_pedido['Valor Total'] = match.group(1)
            
            # Detectar início da seção de produtos
            if 'Cod Forn' in linha_limpa and 'Seq' in linha_limpa and 'Produtos' in linha_limpa:
                in_produtos_section = True
                continue
            
            # Detectar fim da seção de produtos
            # Só considera fim se tiver substantivos além de "TOTAIS" ou se for claramente fim
            if in_produtos_section and ('TOTAIS' in linha_limpa and re.search(r'\d+,\d+', linha_limpa)):
                # Salvar produto atual antes de encerrar (último produto da seção)
                if current_produto and current_produto.get('Código Fornecedor'):
                    produtos_list.append(current_produto.copy())
                in_produtos_section = False
                current_produto = None
                continue
            
            if 'DADOS ADICIONAIS' in linha_limpa or 'ADVERT' in linha_limpa:
                # Salvar produto atual antes de encerrar
                if current_produto and current_produto.get('Código Fornecedor'):
                    produtos_list.append(current_produto.copy())
                in_produtos_section = False
                current_produto = None
                continue
            
            # Processar linhas de produtos
            if in_produtos_section and linha_limpa:
                # Ignorar linhas de cabeçalho da tabela
                if any(word in linha_limpa for word in ['a Receber', 'Fin.', '(Tot.)', '(Unt.)']):
                    continue
                
                # Formato real: ... CodForn(6dig) ... ValorBruto ValorUnit Qtde Emb Seq Descrição
                # Ex: 0,00 6,97 41,82 0,00 0,00 0,00 504251 0,00 0,00 0,00 0,00 41,82 6,9700 6,00 UN 1 CONDIC...
                # Capturar: código + 3 valores finais (bruto, unit, qtde) + emb + seq + descrição
                match_produto = re.search(r'(\d{6})\s+[\d,\.\s]+\s+([\d,\.]+)\s+([\d,\.]+)\s+([\d,\.]+)\s+(UN|CX|PC|KG|LT)\s+(\d+)\s+(.+)$', linha_limpa)
                
                if match_produto:
                    # Salvar produto anterior se existir
                    if current_produto and current_produto.get('Código Fornecedor'):
                        produtos_list.append(current_produto.copy())
                    
                    codigo = match_produto.group(1)  # 6 dígitos
                    valor_bruto = match_produto.group(2)  # Valor bruto (não usado por enquanto)
                    valor_unit = match_produto.group(3)  # Valor unitário
                    qtde = match_produto.group(4)  # Quantidade
                    embalagem = match_produto.group(5)  # UN, CX, etc
                    seq = match_produto.group(6)  # Sequência
                    descricao = match_produto.group(7).strip()  # Descrição
                    
                    current_produto = {
                        'Número do Pedido': current_pedido['Número do Pedido'],
                        'Código Fornecedor': codigo,
                        'Sequência': seq,
                        'Descrição': descricao,
                        'Embalagem': embalagem,
                        'Quantidade': qtde,
                        'Valor Unit.': valor_unit,
                        'EAN': ''
                    }
                elif current_produto:
                    # Linhas adicionais (descrição ou EAN)
                    if 'EAN' in linha_limpa:
                        match_ean = re.search(r'EANs?:\s*([\d,\s]+)', linha_limpa)
                        if match_ean:
                            current_produto['EAN'] = match_ean.group(1).strip()
                    elif linha_limpa and not linha_limpa.startswith('---') and not linha_limpa.startswith('TOTVS'):
                        # Adicionar à descrição se não for linha de sistema ou cabeçalho
                        # Filtrar linhas que parecem ser do cabeçalho (contém CNPJ, Cidade, Bairro, etc)
                        palavras_cabecalho = ['Bairro', 'Cidade', 'CNPJ', 'Endereço', 'Telefone', 'Cep', 'Inscrição']
                        if not any(palavra in linha_limpa for palavra in palavras_cabecalho):
                            if len(linha_limpa) < 50 and not re.match(r'^\d', linha_limpa):
                                current_produto['Descrição'] += ' ' + linha_limpa
        
        # Salvar último pedido e produto
        if current_produto and current_produto.get('Código Fornecedor'):
            produtos_list.append(current_produto)
        if current_pedido:
            pedidos.append(current_pedido)
        
        # Criar DataFrames
        dfs = []
        
        if pedidos:
            df_pedidos = pd.DataFrame(pedidos)
            dfs.append(df_pedidos)
        
        if produtos_list:
            df_produtos = pd.DataFrame(produtos_list)
            
            # Converter colunas numéricas para formato nativo (para o Excel reconhecer como número)
            if not df_produtos.empty:
                def conv_num(val):
                    if isinstance(val, str):
                        # Remove pontos de milhar e troca vírgula decimal por ponto
                        val_clean = val.replace('.', '').replace(',', '.')
                        try:
                            return float(val_clean)
                        except ValueError:
                            return 0.0
                    return val

                # Aplicar conversão em colunas de valores (float)
                cols_valor = ['Quantidade', 'Valor Unit.']
                for col in cols_valor:
                    if col in df_produtos.columns:
                        df_produtos[col] = df_produtos[col].apply(conv_num)
                
                # Converter Identificadores para inteiro
                cols_int = ['Código Fornecedor', 'EAN']
                for col in cols_int:
                    if col in df_produtos.columns:
                        # Remove caracteres não numéricos antes de converter (ex: EANs múltiplos separados por vírgula)
                        # Se houver múltiplos (ex: "789, 790"), pega só o primeiro numérico ou tenta limpar
                        # Aqui vamos usar pd.to_numeric com coerce, que transforma erro em NaN -> 0
                        df_produtos[col] = pd.to_numeric(df_produtos[col], errors='coerce').fillna(0).astype('int64')
                
                # Remover coluna Sequência se existir (não é necessária na saída final)
                if 'Sequência' in df_produtos.columns:
                    df_produtos = df_produtos.drop(columns=['Sequência'])
            
            dfs.append(df_produtos)
        
        return dfs
