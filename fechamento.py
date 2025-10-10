import pandas as pd
import numpy as np
import locale
import chardet

# Configurar locale para português Brasil (usar vírgula como decimal)
try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
except:
    try:
        locale.setlocale(locale.LC_ALL, 'Portuguese_Brazil.1252')
    except:
        print("Não foi possível configurar locale, usando formatação manual")

# Função para detectar a codificação do arquivo
def detectar_codificacao(arquivo_path):
    try:
        with open(arquivo_path, 'rb') as f:
            raw_data = f.read(10000)  # Ler apenas os primeiros 10KB para detecção
            resultado = chardet.detect(raw_data)
            encoding = resultado['encoding']
            confidence = resultado['confidence']
            print(f"Codificação detectada: {encoding} (confiança: {confidence:.2f})")
            return encoding
    except Exception as e:
        print(f"Erro ao detectar codificação: {e}")
        return 'latin-1'  # Fallback para latin-1

# Carregar os arquivos
fechamento_path = r"S:\hor\excel\fechamento-20251001-20251009.csv"
movimentacao_path = r"S:\hor\excel\20251001.csv"

# Detectar codificação dos arquivos
print("Detectando codificação dos arquivos...")
encoding_fechamento = detectar_codificacao(fechamento_path)
encoding_movimentacao = detectar_codificacao(movimentacao_path)

# Ler os arquivos CSV corrigindo o separador e codificação
print("Lendo arquivos...")

def carregar_csv_com_tentativas(caminho, nome_arquivo):
    codificacoes = [encoding_fechamento, 'latin-1', 'ISO-8859-1', 'cp1252', 'utf-8']
    separadores = [';', ',']
    
    for encoding in codificacoes:
        for sep in separadores:
            try:
                print(f"Tentando {nome_arquivo} com encoding={encoding}, sep='{sep}'")
                df = pd.read_csv(caminho, encoding=encoding, sep=sep, decimal=',', on_bad_lines='skip')
                print(f"✅ {nome_arquivo} carregado com encoding={encoding}, sep='{sep}'")
                return df
            except UnicodeDecodeError:
                continue
            except pd.errors.ParserError:
                continue
            except Exception as e:
                print(f"❌ Falha com encoding={encoding}, sep='{sep}': {e}")
                continue
    
    # Se todas as tentativas falharem, tentar método manual
    print(f"Tentando método manual para {nome_arquivo}...")
    try:
        with open(caminho, 'r', encoding='latin-1') as f:
            lines = f.readlines()
        
        if not lines:
            return pd.DataFrame()
            
        # Tentar detectar separador pela primeira linha
        primeira_linha = lines[0].strip()
        if ';' in primeira_linha:
            sep = ';'
        elif ',' in primeira_linha:
            sep = ','
        else:
            sep = ';'  # padrão
            
        # Extrair cabeçalho e dados
        header = [col.strip() for col in primeira_linha.split(sep)]
        data = []
        
        for line in lines[1:]:
            if line.strip():
                # Limpar e processar a linha
                cleaned_line = line.strip()
                # Substituir ponto por vírgula nos números decimais se necessário
                if sep == ';':
                    parts = cleaned_line.split(sep)
                    cleaned_parts = []
                    for part in parts:
                        # Se tem ponto mas não tem vírgula, pode ser decimal com ponto
                        if '.' in part and ',' not in part:
                            part = part.replace('.', ',')
                        cleaned_parts.append(part)
                    data.append(cleaned_parts)
                else:
                    data.append(cleaned_line.split(sep))
        
        # Garantir que todas as linhas tenham o mesmo número de colunas
        max_cols = len(header)
        data_clean = []
        for row in data:
            if len(row) == max_cols:
                data_clean.append(row)
            elif len(row) > max_cols:
                data_clean.append(row[:max_cols])
            else:
                # Preencher com valores vazios se faltarem colunas
                row.extend([''] * (max_cols - len(row)))
                data_clean.append(row)
                
        df = pd.DataFrame(data_clean, columns=header)
        print(f"✅ {nome_arquivo} carregado manualmente")
        return df
        
    except Exception as e:
        print(f"❌ Erro crítico ao carregar {nome_arquivo}: {e}")
        return pd.DataFrame()

# Carregar os arquivos
fechamento_df = carregar_csv_com_tentativas(fechamento_path, "Fechamento")
movimentacao_df = carregar_csv_com_tentativas(movimentacao_path, "Movimentação")

# Verificar se os DataFrames foram carregados
if fechamento_df.empty:
    print("❌ Não foi possível carregar o arquivo de fechamento. Verifique o arquivo.")
    exit(1)

if movimentacao_df.empty:
    print("⚠️  Não foi possível carregar o arquivo de movimentação. Continuando sem ele...")

# Normalizar nomes das colunas (remover espaços extras)
fechamento_df.columns = [col.strip() for col in fechamento_df.columns]
if not movimentacao_df.empty:
    movimentacao_df.columns = [col.strip() for col in movimentacao_df.columns]

print(f"\nColunas no fechamento.csv: {list(fechamento_df.columns)}")
print(f"Total de linhas no fechamento: {len(fechamento_df)}")

# Função para converter valores com vírgula para float
def converter_para_float(valor):
    try:
        if pd.isna(valor) or valor == '':
            return 0.0
        # Se já for número, retorna como float
        if isinstance(valor, (int, float)):
            return float(valor)
        # Se for string, trata vírgula como decimal
        valor_str = str(valor).strip()
        # Remove possíveis pontos de milhar e substitui vírgula por ponto
        valor_str = valor_str.replace('.', '').replace(',', '.')
        return float(valor_str)
    except (ValueError, TypeError):
        return 0.0

# Função CORRIGIDA para calcular Desconto verificado
def calcular_desconto_verificado(grupo, desconto):
    try:
        # Converter para string e limpar
        grupo_str = str(grupo).strip().upper() if pd.notna(grupo) else ""
        
        if grupo_str == "AKKI":
            return 0.03
        elif grupo_str == "ROSSI":
            return 0.05
        elif grupo_str == "TENDA":
            return 0
        else:
            # Para outros grupos, usar o valor da coluna DESCONTO dividido por 100
            try:
                # Converter desconto para float (tratando vírgula como decimal)
                desconto_val = converter_para_float(desconto)
                return desconto_val / 100
            except (ValueError, TypeError):
                return 0
    except:
        return 0

# Verificar os valores únicos na coluna GRUPO para debug
print("\nValores únicos na coluna GRUPO:")
print(fechamento_df['GRUPO'].unique() if 'GRUPO' in fechamento_df.columns else "Coluna GRUPO não encontrada")

# Verificar alguns exemplos de DESCONTO para debug
if 'DESCONTO' in fechamento_df.columns:
    print("\nAmostra dos valores de DESCONTO (originais):")
    print(fechamento_df[['GRUPO', 'DESCONTO']].head(10) if 'GRUPO' in fechamento_df.columns else fechamento_df[['DESCONTO']].head(10))

# Adicionar coluna Desconto verificado CORRIGIDA
print("\nAdicionando coluna 'Desconto verificado'...")
if 'GRUPO' in fechamento_df.columns and 'DESCONTO' in fechamento_df.columns:
    fechamento_df['Desconto verificado'] = fechamento_df.apply(
        lambda row: calcular_desconto_verificado(row.get('GRUPO', ''), row.get('DESCONTO', 0)), 
        axis=1
    )
else:
    print("❌ Colunas GRUPO ou DESCONTO não encontradas para calcular desconto verificado")
    fechamento_df['Desconto verificado'] = 0

# Adicionar coluna Mov (comparação apenas por NOTA FISCAL/NF-E)
print("Adicionando coluna 'Mov'...")

if not movimentacao_df.empty:
    # Verificar se as colunas necessárias existem
    colunas_necessarias_mov = ['NOTA FISCAL', 'DESCRICAO']
    colunas_existem_mov = all(col in movimentacao_df.columns for col in colunas_necessarias_mov)

    if colunas_existem_mov and 'NF-E' in fechamento_df.columns:
        # Criar dicionário para mapeamento rápido
        mov_dict = movimentacao_df.drop_duplicates('NOTA FISCAL').set_index('NOTA FISCAL')['DESCRICAO'].to_dict()
        
        # Preencher coluna Mov
        fechamento_df['Mov'] = fechamento_df['NF-E'].map(mov_dict)
        print("Coluna 'Mov' adicionada com sucesso")
    else:
        fechamento_df['Mov'] = None
        print("Colunas necessárias para 'Mov' não encontradas")
else:
    fechamento_df['Mov'] = None
    print("Arquivo de movimentação não carregado, coluna 'Mov' não adicionada")

# Adicionar coluna Mov V2 (comparação tripla)
print("Adicionando coluna 'Mov V2'...")

if not movimentacao_df.empty:
    # Verificar se as colunas necessárias existem
    colunas_necessarias_v2 = ['NOTA FISCAL', 'ROMANEIO', 'PRODUTO', 'DESCRICAO']
    colunas_existem_v2 = all(col in movimentacao_df.columns for col in colunas_necessarias_v2)
    colunas_existem_fechamento = all(col in fechamento_df.columns for col in ['NF-E', 'ROMANEIO', 'CODPRODUTO'])

    if colunas_existem_v2 and colunas_existem_fechamento:
        # Criar chave composta para comparação mais eficiente
        movimentacao_df['chave_composta'] = (
            movimentacao_df['NOTA FISCAL'].astype(str).str.strip() + '_' + 
            movimentacao_df['ROMANEIO'].astype(str).str.strip() + '_' + 
            movimentacao_df['PRODUTO'].astype(str).str.strip()
        )

        # Criar dicionário para Mov V2
        mov_v2_dict = movimentacao_df.drop_duplicates('chave_composta').set_index('chave_composta')['DESCRICAO'].to_dict()

        # Criar chave composta no fechamento_df para comparação
        fechamento_df['chave_composta'] = (
            fechamento_df['NF-E'].astype(str).str.strip() + '_' + 
            fechamento_df['ROMANEIO'].astype(str).str.strip() + '_' + 
            fechamento_df['CODPRODUTO'].astype(str).str.strip()
        )

        # Preencher coluna Mov V2
        fechamento_df['Mov V2'] = fechamento_df['chave_composta'].map(mov_v2_dict)
        
        # Remover coluna auxiliar
        fechamento_df.drop('chave_composta', axis=1, inplace=True)
        print("Coluna 'Mov V2' adicionada com sucesso")
    else:
        fechamento_df['Mov V2'] = None
        print("Colunas necessárias para 'Mov V2' não encontradas")
else:
    fechamento_df['Mov V2'] = None
    print("Arquivo de movimentação não carregado, coluna 'Mov V2' não adicionada")

# Função para formatar números com vírgula para Excel
def formatar_para_excel(valor):
    try:
        if pd.isna(valor):
            return ""
        # Converter para float primeiro
        valor_float = float(valor)
        # Formatando com vírgula como separador decimal
        return f"{valor_float:.10f}".rstrip('0').rstrip('.').replace('.', ',')
    except:
        return str(valor)

# Aplicar formatação para todas as colunas numéricas (especialmente Desconto verificado)
print("Formatando números para Excel (vírgula como decimal)...")

# Lista de colunas que provavelmente contêm números
colunas_numericas = ['Desconto verificado', 'DESCONTO', 'QTDE', 'QTDE REAL', 'CUSTO', 'FRETE', 
                     'PRODUCAO', 'ESCRITORIO', 'P.COM', 'VLR PIS', 'VLR COFINS', 'IRPJ', 'CSLL',
                     'VLR ICMS', 'ALIQ ICMS', 'VLR DESCONTO', 'PRECO VENDA', 'FAT LIQUIDO',
                     'FAT BRUTO', 'LUCRO', 'MARGEM']

for coluna in colunas_numericas:
    if coluna in fechamento_df.columns:
        fechamento_df[coluna] = fechamento_df[coluna].apply(formatar_para_excel)

# Salvar o novo arquivo com formatação para Excel
output_path = r"C:\Users\win11\Downloads\fechamento_processado.csv"

# Configurar opções para salvar com vírgula como decimal
try:
    fechamento_df.to_csv(output_path, index=False, encoding='utf-8', sep=';', decimal=',')
    print(f"\n✅ Processamento concluído!")
    print(f"Arquivo salvo em: {output_path}")
    print(f"Colunas adicionadas: Desconto verificado, Mov, Mov V2")
    print(f"Total de linhas processadas: {len(fechamento_df)}")
    
    # Mostrar amostra dos resultados com foco no Desconto verificado
    if 'Desconto verificado' in fechamento_df.columns:
        print(f"\nAmostra dos resultados - Desconto verificado:")
        colunas_mostrar = []
        if 'GRUPO' in fechamento_df.columns:
            colunas_mostrar.append('GRUPO')
        if 'DESCONTO' in fechamento_df.columns:
            colunas_mostrar.append('DESCONTO')
        colunas_mostrar.append('Desconto verificado')
        
        print(fechamento_df[colunas_mostrar].head(15))
        
        # Mostrar estatísticas do Desconto verificado
        print(f"\nEstatísticas do Desconto verificado:")
        print(f"Valores únicos: {fechamento_df['Desconto verificado'].unique()}")
        
except Exception as e:
    print(f"❌ Erro ao salvar arquivo: {e}")
    # Tentar salvar com encoding alternativo
    try:
        fechamento_df.to_csv(output_path, index=False, encoding='latin-1', sep=';', decimal=',')
        print(f"✅ Arquivo salvo com encoding latin-1: {output_path}")
    except Exception as e2:
        print(f"❌ Erro crítico ao salvar arquivo: {e2}")