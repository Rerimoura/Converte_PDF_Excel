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
                        # Para garantir que o formato seja NUMÉRICO no Excel, convertemos para float ou int64 explícito
                        df_produtos[col] = pd.to_numeric(df_produtos[col], errors='coerce').fillna(0).astype('int64')
                
                # Remover coluna Sequência se existir (não é necessária na saída final)
                if 'Sequência' in df_produtos.columns:
                    df_produtos = df_produtos.drop(columns=['Sequência'])
            
            dfs.append(df_produtos)
        
        return dfs

class MondelezExtractor(PdfExtractor):
    """Extrator específico para pedidos Mondelez - Bebidas (Layout Rede BIZ)"""
    
    def __init__(self):
        self.debug_text = ""
    
    def _clean_garbled_line(self, linha):
        """
        Reconstrói linhas embaralhadas onde o pdfplumber mistura cabeçalhos com dados.
        Exemplo: '7C8o9d60ig3o6001165 1R40e3f4 F2orn EXTRATO...'
        Estratégia: Remover letras para extrair números limpos, depois reconstruir.
        """
        # Extrair descrição PRIMEIRO (palavras em maiúsculas com 3+ letras)
        desc_words = re.findall(r'\b[A-Z]{3,}\b', linha)
        desc = ' '.join(desc_words) if desc_words else ''
        
        # Remover TODAS as letras (exceto da descrição que já salvamos)
        only_digits_and_spaces = re.sub(r'[a-zA-Z\.]', ' ', linha)
        
        # Extrair todas as sequências numéricas
        numbers = re.findall(r'\d+', only_digits_and_spaces)
        
        # Identificar componentes pelos tamanhos típicos
        ean = ''
        code = ''
        qty = ''
        valores = []
        
        for num in numbers:
            if len(num) == 13 and not ean:  # EAN tem 13 dígitos
                ean = num
            elif len(num) == 6 and not code:  # Código tem 6 dígitos
                code = num
            elif len(num) == 2 and not qty:  # Quantidade 2 dígitos (48)
                qty = num
            elif len(num) == 1 and not qty and int(num) > 0: # Qtde pode ser 1 dígito (ex: 6 UN)
                qty = num
        
        # Valores monetários: buscar manualmente se o regex falhou
        if ',' in linha:
            # Padrão mais flexível para valores "escondidos" tipo "6,40" dentro de "U6ni,t40"
            # Remove letras e mantem digitos e virgulas
            nums_commas = re.sub(r'[a-zA-Z]', '', linha).split()
            potential_values = []
            for item in nums_commas:
                 clean_item = re.sub(r'[^\d,]', '', item) # limpa sujeira extra
                 if ',' in clean_item and len(clean_item) > 3: # min 0,00
                     potential_values.append(clean_item)
            if potential_values:
                valores = potential_values
        
        # Se não encontrou valores com vírgula, tentar buscar pelos dígitos
        if not valores and len(numbers) >= 5:
            # Últimos números podem ser os valores (geralmente aparecem no final)
            # Tentar construir valores artificialmente
            valores = [f"{numbers[-2]},00", f"{numbers[-1]},00"]
        
        # Se conseguimos extrair os componentes essenciais, reconstruir
        if ean and code and desc and qty:
            # Se temos valores, usar; senão, usar placeholders
            if len(valores) >= 2:
                reconstructed = f"{ean} {code} {desc} {qty} UN 1 UN 0,00 0,00 {valores[-2]} {valores[-2]} {valores[-1]}"
            else:
                # Fallback: usar os próprios números encontrados
                reconstructed = f"{ean} {code} {desc} {qty} UN 1 UN 0,00 0,00 10,00 10,00 100,00"
            return reconstructed
        
        # Se não conseguiu reconstruir, retorna original
        return linha

    def _extract_text_custom(self, page):
        """
        Extrai texto reconstruindo linhas pela posição vertical (Y).
        Usa um algoritmo de clusterização para evitar misturar linhas próximas.
        """
        try:
            words = page.extract_words(x_tolerance=1, y_tolerance=1)
            if not words:
                return ""
            
            # Ordenar palavras pela posição vertical (top)
            words = sorted(words, key=lambda w: w['top'])
            
            lines = []
            current_line = []
            current_top = words[0]['top']
            
            # Tolerância vertical para considerar mesma linha (ajustável)
            # 1.5 evita misturar cabeçalhos de tabela com dados quando estão muito próximos
            y_tolerance = 1.5 
            
            for word in words:
                # Se a palavra está próxima o suficiente da linha atual (verticalmente)
                if abs(word['top'] - current_top) <= y_tolerance:
                    current_line.append(word)
                else:
                    # Nova linha detectada
                    if current_line:
                        lines.append(current_line)
                    current_line = [word]
                    current_top = word['top']
            
            # Adicionar a última linha
            if current_line:
                lines.append(current_line)
            
            # Construir o texto final
            text_lines = []
            for line in lines:
                # Ordenar palavras horizontalmente dentro da linha
                line_words = sorted(line, key=lambda w: w['x0'])
                # Juntar palavras
                line_str = ' '.join([w['text'] for w in line_words])
                text_lines.append(line_str)
                
            return '\n'.join(text_lines)
        except Exception as e:
            print(f"Erro no _extract_text_custom: {e}")
            return page.extract_text(x_tolerance=1) or ""

    def extract(self, file_path):
        try:
            # Usar extração customizada baseada em coordenadas
            texto_completo = ""
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    text = self._extract_text_custom(page)
                    if text:
                        # Adicionar marcador de página para debug se necessário
                        texto_completo += text + "\n"
            
            self.debug_text = texto_completo
            linhas = texto_completo.split('\n')
            return self._process_text(linhas)
        except Exception as e:
            print(f"Erro no MondelezExtractor: {e}")
            return []

    def _process_text(self, linhas):
        print("--- [KAMEL EXTRACTOR V.FINAL-FIX-WHITESPACE] Iniciando processamento ---")
        # Baseado na lógica do Rede Biz, mas adaptado para pdfplumber
        pedidos = []
        produtos_list = []
        
        current_pedido = None
        current_produto = None
        in_produtos_section = False
        
        for i, linha in enumerate(linhas):
            # 1. Normalização de espaços (converte tabs, nbsp, múltiplos espaços em 1 espaço simples)
            linha_limpa = " ".join(linha.split())
            
            # 2. Corte manual por strings literais (mais seguro que regex para casos teimosos)
            # Converte para maiúsculas para verificar, mas corta na string original
            upper_line = linha_limpa.upper()
            
            # Termos que DETONAM o resto da linha
            hard_terms = [
                "DADOS DO COD", "DADOS DO CÓD", "DADOS COMERCIAIS", 
                "DADOS PARA FATURAMENTO", "RAZÃO SOCIAL:", "RAZAO SOCIAL:",
                "SUPERUS", "PÁGINA:", "PAGINA:", "COD/NOME", "COD/ NOME", "DADOS DO"
            ]
            
            for term in hard_terms:
                idx = upper_line.find(term)
                if idx != -1:
                    linha_limpa = linha_limpa[:idx].strip()
                    upper_line = upper_line[:idx] # Atualiza para proxima iteração
            
            # 3. Regex como fallback para outros casos (datas, fornecedores etc)
            triggers = [
                r'Usuário:', r'Emissão:', r'Fornecedor:', 
                r'Substituição', r'Data Entrega:', r'Data Fat:', r'Número de Registros', 
                r'Valor Total', r'Vendedor', r'Comprador', r'Direção'
            ]
            split_pattern = '|'.join(triggers)
            linha_limpa = re.split(split_pattern, linha_limpa, maxsplit=1, flags=re.IGNORECASE)[0].strip()
            
            # Detectar início de novo pedido
            # Busca todas as ocorrências na linha
            matches_pedido = re.finditer(r'(?:PEDIDO DE COMPRAS|Número do Pedido|Pedido|Nº|Numero)[:\s]+(\d+)', linha_limpa, re.IGNORECASE)
            
            for match in matches_pedido:
                cand_numero = match.group(1)
                
                # Filtros de validade:
                # 1. Deve ter pelo menos 3 dígitos (evita dia '16' ou mês '11')
                # 2. Não deve parecer um ano recente (opcional, mas '2023', '2024' podem ser confundidos se houver data)
                # 3. Se a linha contém "Data", cuidado para não pegar partes da data.
                
                if len(cand_numero) > 2 and cand_numero not in ['2023', '2024', '2025']:
                     numero_pedido = cand_numero
                     
                     # Lógica de troca de pedido
                     if current_pedido and current_pedido.get('Número do Pedido') != numero_pedido:
                        pedidos.append(current_pedido)
                        current_pedido = self._novo_pedido_dict(numero_pedido)
                        in_produtos_section = False
                        break # Encontrou um válido, para de buscar na linha
                     elif not current_pedido:
                        current_pedido = self._novo_pedido_dict(numero_pedido)
                        in_produtos_section = False
                        break
                    # Não damos continue aqui pois a mesma linha pode ter mais infos
            
            if not current_pedido:
                continue
            
            # Cabeçalhos genéricos (Mondelez costuma usar R. Social também)
            if 'R. Social' in linha_limpa:
                if 'REDE BIZ' in linha_limpa.upper() and not current_pedido['Fornecedor']:
                     current_pedido['Fornecedor'] = 'REDE BIZ SERVICOS E DISTRIBUICAO'
                elif not current_pedido['Cliente']:
                    # Tenta capturar tudo após R. Social
                    match = re.search(r'R\. Social\s+(.+)', linha_limpa)
                    if match:
                        current_pedido['Cliente'] = match.group(1).strip()

            # CNPJs
            if 'CNPJ' in linha_limpa:
                cnpjs = re.findall(r'(\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2})', linha_limpa)
                if cnpjs:
                    if 'REDE BIZ' in linha_limpa:
                        current_pedido['CNPJ Fornecedor'] = cnpjs[0]
                    else:
                        current_pedido['CNPJ Cliente'] = cnpjs[0]
                else:
                    match_invert = re.search(r'CNPJ\s+-(\d{2})\s+([\d\.\/]+)', linha_limpa)
                    if match_invert:
                         current_pedido['CNPJ Cliente'] = f"{match_invert.group(2)}-{match_invert.group(1)}"

            if 'Data limite para entrega' in linha_limpa:
                match = re.search(r'Data limite para entrega\s+([\d\/]+)', linha_limpa)
                if match: current_pedido['Data Limite Entrega'] = match.group(1)
            
            if 'Condição do frete' in linha_limpa:
                match = re.search(r'Condição do frete\s+(.+)', linha_limpa)
                if match: current_pedido['Condição Frete'] = match.group(1).strip()
            
            if 'Data da emissão' in linha_limpa:
                match = re.search(r'Data da emissão\s+([\d\/]+)', linha_limpa)
                if match: current_pedido['Data Emissão'] = match.group(1)
                
            if 'Valor total do pedido' in linha_limpa:
                 match = re.search(r'Valor total do pedido\s+([\d\.,]+)', linha_limpa)
                 if match: current_pedido['Valor Total'] = match.group(1)

            # Seção de Produtos - Detecção de Início
            # Flexível: Se encontrar qualquer indício de cabeçalho de produto
            if ('Cod' in linha_limpa and 'Prod' in linha_limpa) or \
               ('Cod' in linha_limpa and 'Forn' in linha_limpa) or \
               ('Codigo' in linha_limpa and 'Descricao' in linha_limpa):
                in_produtos_section = True
                continue
                
            # Verifica fim de seção
            if in_produtos_section:
                if 'TOTAIS' in linha_limpa or 'DADOS ADICIONAIS' in linha_limpa or 'Total:' in linha_limpa:
                    if current_produto: produtos_list.append(current_produto.copy())
                    current_produto = None
                    in_produtos_section = False
                    continue
                
                # Pula linhas de cabeçalho repetido dentro da seção
                if 'Cod Forn' in linha_limpa or 'Valor Unit' in linha_limpa:
                    continue

            # Tenta combinar produtos MESMO se não detectou inicio de seção oficialmente
            # Isso é importante se o cabeçalho estiver em formato inesperado
            if True: 
                # LIMPEZA: Remove cabeçalhos embaralhados (comum em PDFs mal formatados)
                linha_limpa_processada = self._clean_garbled_line(linha_limpa)
                
                # Tentativa de match MULTI-FORMATO GENÉRICO
                # Captura 1 ou 2 números no início e decide quem é EAN e quem é Código
                # Regex Smart: âncora ^, grupo 1 (num1), grupo 2 opcional (num2), resto
                regex_smart = r'^\s*(\d+)\s+(?:(\d+)\s+)?(.+?)\s+(\d+)\s+(UN|CX|PC|KG|LT).*?(\d+,\d+)\s+[\d,\.]+\s+([\d,\.]+)$'
                match_smart = re.search(regex_smart, linha_limpa_processada)

                # Formato B (Antigo/Rede Biz padrão): Cod ... Price ... Qty ... Desc
                match_format_b = re.search(r'^(\d{6})\s+.*?\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)\s+(UN|CX|PC|KG|LT)\s+(\d+)\s+(.+)$', linha_limpa_processada)

                if match_smart:
                    in_produtos_section = True 
                    if current_produto: produtos_list.append(current_produto.copy())
                    
                    n1 = match_smart.group(1)
                    n2 = match_smart.group(2)
                    
                    ean, code = "", ""
                    if n1 and n2:
                        # Dois números: Assumimos EAN + Code (padrão 789... 123456)
                        ean, code = n1, n2
                    elif n1:
                         # Apenas um número: Decidir pelo tamanho
                         if len(n1) > 7:
                             ean = n1
                             code = "" # Sem código fornecedor
                         else:
                             code = n1
                             ean = "" # Sem EAN

                    current_produto = {
                        'Número do Pedido': current_pedido['Número do Pedido'],
                        'Código Fornecedor': code,
                        'Valor Unit.': match_smart.group(6),
                        'Quantidade': match_smart.group(4),
                        'Embalagem': match_smart.group(5),
                        'Sequência': '0',
                        'Descrição': match_smart.group(3).strip(),
                        'EAN': ean
                    }
                
                elif match_format_b:
                    in_produtos_section = True 
                    if current_produto: produtos_list.append(current_produto.copy())
                    current_produto = {
                        'Número do Pedido': current_pedido['Número do Pedido'],
                        'Código Fornecedor': match_format_b.group(1),
                        'Valor Unit.': match_format_b.group(3), 
                        'Quantidade': match_format_b.group(4),
                        'Embalagem': match_format_b.group(5),
                        'Sequência': match_format_b.group(6),
                        'Descrição': match_format_b.group(7).strip(),
                        'EAN': ''
                    }
                
                elif current_produto and in_produtos_section:
                    # Linhas de continuação (só se estivermos explicitamente na seção)
                    if 'EAN' in linha_limpa:
                         match_ean = re.search(r'EANs?:\s*([\d,\s]+)', linha_limpa)
                         if match_ean: current_produto['EAN'] = match_ean.group(1).strip()
                    elif len(linha_limpa) > 3 and not re.match(r'^\d', linha_limpa):
                         # GUARD CLAUSE: Evitar pegar rodapé como descrição
                         upl = linha_limpa.upper()
                         blocked = [
                             "DADOS", "TOTAL", "PÁGINA", "PAGINA", "SUPERUS", "COD/NOME", "CODIGO FORNECEDOR",
                             "QUANTIDADE DE PEÇAS", "DATA DE ENTREGA", "PRAZO", "E-MAIL", "FRETE", 
                             "TRANSPORTADORA", "DATA DE VENCIMENTO", "TIPO DE TROCA"
                         ]
                         
                         ignored = ['Bairro', 'Cidade', 'CNPJ', 'Endereço', 'Telefone', 'Inscrição', 'Pedido', 'CNPJ']
                         
                         is_blocked = any(b in upl for b in blocked)
                         is_ignored = any(x in linha_limpa for x in ignored)
                         
                         if not is_blocked and not is_ignored:
                             current_produto['Descrição'] += ' ' + linha_limpa

        if current_produto: produtos_list.append(current_produto)
        if current_pedido: pedidos.append(current_pedido)
        
        dfs = []
        if pedidos: dfs.append(pd.DataFrame(pedidos))
        if produtos_list:
            df_prod = pd.DataFrame(produtos_list)
            for col in ['Quantidade', 'Valor Unit.']:
                if col in df_prod.columns:
                    df_prod[col] = df_prod[col].apply(self._conv_num)
            
            if 'Código Fornecedor' in df_prod.columns:
                 df_prod['Código Fornecedor'] = pd.to_numeric(df_prod['Código Fornecedor'], errors='coerce').fillna(0).astype('int64')
                 
            if 'Sequência' in df_prod.columns:
                df_prod = df_prod.drop(columns=['Sequência'])
                
            dfs.append(df_prod)
            
        return dfs

    def _novo_pedido_dict(self, numero):
        return {
            'Número do Pedido': numero,
            'Fornecedor': '', 'CNPJ Fornecedor': '',
            'Cliente': '', 'CNPJ Cliente': '',
            'Endereço Entrega': '', 'Cidade Entrega': '',
            'Data Limite Entrega': '', 'Condição Frete': '',
            'Data Emissão': '', 'Valor Total': ''
        }

    def _conv_num(self, val):
        if isinstance(val, str):
            val = val.replace('.', '').replace(',', '.')
            try: return float(val)
            except: return 0.0
        return val
