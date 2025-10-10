import pandas as pd
import numpy as np
from datetime import date
import warnings
import json
import openpyxl.styles
from openpyxl.styles import Alignment

warnings.filterwarnings('ignore')

print("🚀 INICIANDO PROCESSAMENTO DE MARGEM...")

# Data fixa para o nome do arquivo
data_nome_arquivo = "testes"

# Função para carregar CSV com múltiplas tentativas de codificação
def carregar_csv_com_codificacao(caminho, sep=';', decimal=',', thousands='.', skiprows=None):
    codificacoes = ['latin-1', 'iso-8859-1', 'cp1252', 'utf-8']
    
    for encoding in codificacoes:
        try:
            if skiprows:
                df = pd.read_csv(caminho, sep=sep, encoding=encoding, decimal=decimal, thousands=thousands, skiprows=skiprows)
            else:
                df = pd.read_csv(caminho, sep=sep, encoding=encoding, decimal=decimal, thousands=thousands)
            print(f"✅ Arquivo carregado com {encoding}")
            return df
        except UnicodeDecodeError as e:
            print(f"❌ Falha com {encoding}: {e}")
            continue
        except Exception as e:
            print(f"❌ Erro com {encoding}: {e}")
            continue
    
    # Última tentativa com tratamento de erros
    try:
        print("Tentando com tratamento de erros...")
        if skiprows:
            df = pd.read_csv(caminho, sep=sep, encoding='utf-8', decimal=decimal, thousands=thousands, 
                           skiprows=skiprows, errors='replace')
        else:
            df = pd.read_csv(caminho, sep=sep, encoding='utf-8', decimal=decimal, thousands=thousands, 
                           errors='replace')
        print("✅ Arquivo carregado com substituição de caracteres")
        return df
    except Exception as e:
        print(f"❌ Erro crítico ao carregar {caminho}: {e}")
        return pd.DataFrame()

# Carregar arquivos
print("📁 Carregando arquivos...")

# Carregar fechamento
fechamento = carregar_csv_com_codificacao(r"C:\Users\win11\Downloads\fechamento_processado.csv")

# Carregar cancelados (com skiprows=2)
cancelados = carregar_csv_com_codificacao(r"S:\hor\arquivos\gustavo\can.csv", skiprows=2)

# Carregar devoluções (arquivo onde vamos buscar o PESO para QTDE AJUSTADA)
devolucoes = carregar_csv_com_codificacao(r"S:\hor\excel\20251001.csv")

# Carregar custos_produtos (Excel)
try:
    custos_produtos = pd.read_excel(r"C:\Users\win11\Downloads\Custos de produtos - Outubro.xlsx", sheet_name='Base', dtype=str)
    print("✅ Custos produtos carregado com sucesso")
except Exception as e:
    print(f"❌ Erro ao carregar custos produtos: {e}")
    custos_produtos = pd.DataFrame()

# Carregar LOURENCINI
try:
    lourencini = pd.read_excel(r"C:\Users\win11\Downloads\LOURENCINI.xlsx")
    required_cols = ['COD', '0,2', '0,3', '0,5', '0,7', '1', 'Data']
    if all(col in lourencini.columns for col in required_cols):
        lourencini['COD'] = lourencini['COD'].astype(str).str.strip()
        lourencini['COD'] = lourencini['COD'].str.replace(r'\.0$', '', regex=True)
        lourencini['COD'] = lourencini['COD'].str.replace(r'^0+', '', regex=True)
        lourencini['COD'] = lourencini['COD'].str.strip()
        lourencini = lourencini[lourencini['COD'] != '']
        lourencini = lourencini[lourencini['COD'] != 'nan']
        lourencini = lourencini.dropna(subset=['COD'])
        
        # Converter COD para inteiro onde possível
        def converter_para_int_se_possivel(valor):
            try:
                if pd.isna(valor) or valor == '':
                    return np.nan
                return int(float(valor))
            except (ValueError, TypeError):
                return np.nan
        
        lourencini['COD'] = lourencini['COD'].apply(converter_para_int_se_possivel)
        lourencini = lourencini.dropna(subset=['COD'])
        
        colunas_preco = ['0,2', '0,3', '0,5', '0,7', '1']
        for col in colunas_preco:
            lourencini[col] = pd.to_numeric(lourencini[col], errors='coerce')
        
        lourencini['Data'] = pd.to_datetime(lourencini['Data'], errors='coerce', dayfirst=True)
        if 'Data_fim' in lourencini.columns:
            lourencini['Data_fim'] = pd.to_datetime(lourencini['Data_fim'], errors='coerce', dayfirst=True)
        
        lourencini = lourencini.sort_values('Data')
        print("✅ LOURENCINI carregado com sucesso")
    else:
        lourencini = pd.DataFrame()
        print("⚠️ Colunas necessárias não encontradas no LOURENCINI")
except Exception as e:
    lourencini = pd.DataFrame()
    print(f"⚠️ Erro ao carregar LOURENCINI: {e}")

# Carregar OFERTAS_VOG
try:
    # Tentar carregar ambas as abas
    ofertas_off = pd.read_excel(r"C:\Users\win11\Downloads\OFERTAS_VOG.xlsx", sheet_name='OFF_VOG')
    ofertas_cb = pd.read_excel(r"C:\Users\win11\Downloads\OFERTAS_VOG.xlsx", sheet_name='OFF_VOG_CB')
    
    # Combinar as duas abas em um único dataframe
    ofertas_vog = pd.concat([ofertas_off, ofertas_cb], ignore_index=True)
    print("✅ OFERTAS_VOG carregado com sucesso (ambas as abas)")
except Exception as e:
    print(f"⚠️ Erro ao carregar OFERTAS_VOG: {e}")
    ofertas_vog = pd.DataFrame()

# Verificar se os DataFrames essenciais foram carregados
if fechamento.empty:
    print("❌ CRÍTICO: DataFrame 'fechamento' está vazio! Verifique o arquivo.")
    exit()

if cancelados.empty:
    print("⚠️ AVISO: DataFrame 'cancelados' está vazio!")

if devolucoes.empty:
    print("⚠️ AVISO: DataFrame 'devolucoes' está vazio!")

print("✅ TODOS OS ARQUIVOS CARREGADOS COM SUCESSO!")
print(f"📊 Tamanho do fechamento: {len(fechamento)} linhas")
print(f"📊 Tamanho do cancelados: {len(cancelados)} linhas")
print(f"📊 Tamanho do devoluções: {len(devolucoes)} linhas")

# Resto do código original continua aqui...
# [INSIRA AQUI O RESTO DO SEU CÓDIGO ORIGINAL A PARTIR DA LINHA ONDE COMEÇA "Renomear colunas"]

# Renomear colunas
rename_mapping = {
    'PRODUTO': 'CODPRODUTO', 'DATA': 'DATA', 'PCS': 'QTDE', 'KGS': 'PESO_KGS', 
    'CUSTO': 'CUSTO', 'FRETE': 'FRETE', 'PRODUÇÃO': 'PRODUÇÃO', 'TOTAL': 'TOTAL', 
    'QTD': 'QTD', 'PESO': 'PESO'
}
custos_produtos = custos_produtos.rename(columns=rename_mapping)

# Converter colunas numéricas
colunas_numericas = ['PESO_KGS', 'CUSTO', 'FRETE', 'PRODUÇÃO', 'TOTAL', 'QTD', 'PESO']
for coluna in colunas_numericas:
    if coluna in custos_produtos.columns:
        try:
            if custos_produtos[coluna].dtype != 'object':
                custos_produtos[coluna] = custos_produtos[coluna].astype(str)
            custos_produtos[coluna] = custos_produtos[coluna].apply(
                lambda x: str(x).replace(',', '.').replace(' ', '') if pd.notna(x) else x
            )
            custos_produtos[coluna] = pd.to_numeric(custos_produtos[coluna], errors='coerce')
        except Exception:
            pass

custos_produtos['DATA'] = pd.to_datetime(custos_produtos['DATA'], errors='coerce', dayfirst=True)

# Função para converter CODPRODUTO para inteiro
def converter_codproduto_para_int(df, coluna='CODPRODUTO'):
    df[coluna] = df[coluna].astype(str).str.strip()
    df[coluna] = df[coluna].str.replace(r'\.0$', '', regex=True)
    df[coluna] = df[coluna].str.replace(r'^0+', '', regex=True)
    df[coluna] = df[coluna].str.strip()
    
    def converter_para_int(valor):
        try:
            if pd.isna(valor) or valor == '' or valor == 'nan':
                return np.nan
            # Tentar converter para float primeiro e depois para int
            return int(float(valor))
        except (ValueError, TypeError):
            return np.nan
    
    df[coluna] = df[coluna].apply(converter_para_int)
    return df

custos_produtos = converter_codproduto_para_int(custos_produtos)

# Processar dados
print("🔄 Processando dados...")
notas_canceladas = cancelados['NUMERO'].tolist() if not cancelados.empty else []

if not devolucoes.empty:
    devolucoes_filtro = devolucoes[
        (devolucoes['DESCRICAO'] == "DEV VENDA C/ FIN S/ EST") | 
        (devolucoes['HISTORICO'] == "68")
    ]
    devolucoes_var = devolucoes_filtro[['ROMANEIO', 'NOTA FISCAL']].values.tolist()

    vendas_filtro = devolucoes[
        (devolucoes['DESCRICAO'] == "VENDA") | 
        (devolucoes['HISTORICO'] == "51")
    ]
    vendas_var = vendas_filtro[['ROMANEIO', 'NOTA FISCAL']].values.tolist()
else:
    devolucoes_var = []
    vendas_var = []

fechamento_sem_cancelados = fechamento[~fechamento['NF-E'].isin(notas_canceladas)].copy()

# Criar dicionário de custos
custos_dict = {}
custos_data = custos_produtos.to_dict('records')

for i, row in enumerate(custos_data):
    try:
        data_value = row.get('DATA', None)
        codproduto_value = row.get('CODPRODUTO', None)
        
        if data_value is None or pd.isna(data_value) or (isinstance(data_value, str) and data_value.strip() == ''):
            continue
        if codproduto_value is None or pd.isna(codproduto_value) or (isinstance(codproduto_value, str) and codproduto_value.strip() == ''):
            continue
        
        codproduto = str(codproduto_value).strip()
        data_key = pd.to_datetime(data_value, errors='coerce', dayfirst=True)
        if pd.isna(data_key):
            continue
        data_key = data_key.date()
        
        custo_val = float(row.get('CUSTO', 0)) if pd.notna(row.get('CUSTO')) and str(row.get('CUSTO', '')).strip() != '' else 0
        peso_val = float(row.get('PESO', 1)) if pd.notna(row.get('PESO')) and str(row.get('PESO', '')).strip() != '' else 1
        producao_val = float(row.get('PRODUÇÃO', 0)) if pd.notna(row.get('PRODUÇÃO')) and str(row.get('PRODUÇÃO', '')).strip() != '' else 0
        frete_val = float(row.get('FRETE', 0)) if pd.notna(row.get('FRETE')) and str(row.get('FRETE', '')).strip() != '' else 0
        qtd_val = float(row.get('QTD', 0)) if pd.notna(row.get('QTD')) and str(row.get('QTD', '')).strip() != '' else 0
        
        custos_dict[(codproduto, data_key)] = {
            'QTD': qtd_val, 'PESO': peso_val, 'CUSTO': custo_val, 
            'FRETE': frete_val, 'PRODUÇÃO': producao_val
        }
    except Exception:
        continue

# Dicionários para lookup
quinzena_dict = {}
fechamento['PK'] = fechamento['ROMANEIO'].astype(str) + "_" + fechamento['NF-E'].astype(str) + "_" + fechamento['CODPRODUTO'].astype(str)
for _, row in fechamento.iterrows():
    try:
        if pd.notna(row['QUINZENA']):
            quinzena_value = str(row['QUINZENA'])
            # Converter para o novo formato
            if quinzena_value == "Primeira Quinzena":
                quinzena_dict[row['PK']] = "1ª Quinzena"
            elif quinzena_value == "Segunda Quinzena":
                quinzena_dict[row['PK']] = "2ª Quinzena"
            else:
                quinzena_dict[row['PK']] = quinzena_value
    except:
        continue

comissao_regra_dict = {}
for _, row in fechamento.iterrows():
    try:
        key = (int(row['ROMANEIO']), int(row['NF-E']), int(row['CODPRODUTO']))
        comissao_regra_dict[key] = row['P.COM'] if pd.notna(row['P.COM']) else 0
    except:
        continue

fechamento_pk_dict = {}
for _, row in fechamento.iterrows():
    pk = str(row['ROMANEIO']) + "_" + str(row['NF-E']) + "_" + str(row['CODPRODUTO'])
    desconto_verificado = row['Desconto verificado'] if 'Desconto verificado' in fechamento.columns and pd.notna(row['Desconto verificado']) else np.nan
    
    quinzena_value = row['QUINZENA'] if 'QUINZENA' in fechamento.columns and pd.notna(row['QUINZENA']) else ""
    # Converter para o novo formato
    if quinzena_value == "Primeira Quinzena":
        quinzena_value = "1ª Quinzena"
    elif quinzena_value == "Segunda Quinzena":
        quinzena_value = "2ª Quinzena"
    
    fechamento_pk_dict[pk] = {
        'ESCRITORIO': row['ESCRITORIO'] if 'ESCRITORIO' in fechamento.columns and pd.notna(row['ESCRITORIO']) else np.nan,
        'VLR ICMS': row['VLR ICMS'] if 'VLR ICMS' in fechamento.columns and pd.notna(row['VLR ICMS']) else np.nan,
        'PRECO VENDA': row['PRECO VENDA'] if 'PRECO VENDA' in fechamento.columns and pd.notna(row['PRECO VENDA']) else np.nan,
        'QUINZENA': quinzena_value,
        'DESCONTO_VERIFICADO': desconto_verificado,
        'MOV': row['Mov'] if 'Mov' in fechamento.columns and pd.notna(row['Mov']) else "",
        'MOV_V2': row['Mov V2'] if 'Mov V2' in fechamento.columns and pd.notna(row['Mov V2']) else ""
    }

fechamento_nf_dict = {}
for _, row in fechamento.iterrows():
    if pd.notna(row['NF-E']):
        fechamento_nf_dict[int(row['NF-E'])] = row['Mov'] if 'Mov' in fechamento.columns and pd.notna(row['Mov']) else ""

# CRIAR DICIONÁRIO PARA QTDE AJUSTADA A PARTIR DO ARQUIVO DE DEVOLUÇÕES
print("📋 Criando dicionário para QTDE AJUSTADA...")
qtde_ajustada_dict = {}

if not devolucoes.empty:
    # Verificar se as colunas necessárias existem
    colunas_necessarias = ['NOTA FISCAL', 'ROMANEIO', 'PRODUTO', 'PESO']
    colunas_existentes = [col for col in colunas_necessarias if col in devolucoes.columns]
    
    if len(colunas_existentes) == 4:
        # Processar cada linha do arquivo de devoluções
        for _, row in devolucoes.iterrows():
            try:
                nota_fiscal = row['NOTA FISCAL']
                romaneio = row['ROMANEIO']
                produto = row['PRODUTO']
                peso = row['PESO']
                
                # Converter para tipos consistentes (todos string para comparação)
                if pd.notna(nota_fiscal) and pd.notna(romaneio) and pd.notna(produto) and pd.notna(peso):
                    nota_fiscal_str = str(nota_fiscal).strip()
                    romaneio_str = str(romaneio).strip()
                    produto_str = str(produto).strip()
                    
                    # Criar chave única para o dicionário
                    chave = (nota_fiscal_str, romaneio_str, produto_str)
                    
                    # Converter peso para numérico
                    try:
                        peso_float = float(str(peso).replace(',', '.'))
                        qtde_ajustada_dict[chave] = peso_float
                    except (ValueError, TypeError):
                        continue
                        
            except Exception as e:
                continue
        
        print(f"✅ Dicionário QTDE AJUSTADA criado com {len(qtde_ajustada_dict)} entradas")
    else:
        print(f"⚠️ Colunas necessárias não encontradas no arquivo de devoluções. Colunas existentes: {list(devolucoes.columns)}")
        qtde_ajustada_dict = {}
else:
    print("⚠️ Arquivo de devoluções vazio, dicionário QTDE AJUSTADA não criado")
    qtde_ajustada_dict = {}

# Criar base_df
print("📊 Criando base principal...")
base_df = pd.DataFrame()

# Preencher colunas básicas
base_df['CF'] = fechamento_sem_cancelados.apply(
    lambda row: 'DEV' if any([str(row['ROMANEIO']) == str(dev[0]) and str(row['NF-E']) == str(dev[1]) for dev in devolucoes_var]) 
    else row['LOJA'], axis=1
)
base_df['RAZAO'] = fechamento_sem_cancelados['RAZAO']
base_df['FANTASIA'] = fechamento_sem_cancelados['FANTASIA']
base_df['GRUPO'] = fechamento_sem_cancelados['GRUPO']
base_df['OS'] = fechamento_sem_cancelados['ROMANEIO']
base_df['NF-E'] = fechamento_sem_cancelados['NF-E']
base_df['CF_NF'] = fechamento_sem_cancelados['CF_NF'].fillna("")
base_df['DATA'] = pd.to_datetime(fechamento_sem_cancelados['DATA'], dayfirst=True, errors='coerce').dt.date
base_df['VENDEDOR'] = fechamento_sem_cancelados['VENDEDOR']

# Converter CODPRODUTO para inteiro
base_df['CODPRODUTO'] = fechamento_sem_cancelados['CODPRODUTO'].astype(str)
base_df = converter_codproduto_para_int(base_df)

base_df['GRUPO PRODUTO'] = fechamento_sem_cancelados['GRUPO PRODUTO']
base_df['DESCRICAO'] = fechamento_sem_cancelados['DESCRICAO']
base_df['QTDE'] = fechamento_sem_cancelados['QTDE']  
base_df['QTDE REAL'] = fechamento_sem_cancelados['QTDE REAL'] 
base_df['CUSTO EM SISTEMA'] = fechamento_sem_cancelados['CUSTO']
base_df['Val Pis'] = fechamento_sem_cancelados['VLR PIS'].fillna(0) if 'VLR PIS' in fechamento_sem_cancelados.columns else 0
base_df['VLRCOFINS'] = fechamento_sem_cancelados['VLR COFINS'].fillna(0) if 'VLR COFINS' in fechamento_sem_cancelados.columns else 0
base_df['IRPJ'] = fechamento_sem_cancelados['IRPJ'].fillna(0) if 'IRPJ' in fechamento_sem_cancelados.columns else 0
base_df['CSLL'] = fechamento_sem_cancelados['CSLL'].fillna(0) if 'CSLL' in fechamento_sem_cancelados.columns else 0
base_df['VL ICMS'] = fechamento_sem_cancelados['VLR ICMS'] if 'VLR ICMS' in fechamento_sem_cancelados.columns else 0
base_df['Preço Venda'] = fechamento_sem_cancelados['PRECO VENDA'] if 'PRECO VENDA' in fechamento_sem_cancelados.columns else 0

base_df['PK'] = base_df['OS'].astype(str) + "_" + base_df['NF-E'].astype(str) + "_" + base_df['CODPRODUTO'].astype(str)
base_df['Quinzena'] = base_df['PK'].map(lambda x: quinzena_dict.get(x, ""))
base_df['GRUPO'] = base_df['GRUPO'].fillna('VAREJO')    

# FUNÇÃO MODIFICADA PARA CALCULAR QTDE AJUSTADA (COM CONDIÇÃO PARA VALORES NEGATIVOS)
def calcular_qtde_ajustada(row):
    try:
        # Buscar no dicionário usando NF-E, OS e CODPRODUTO
        nf_e = str(row['NF-E']).strip() if pd.notna(row['NF-E']) else ""
        os_val = str(row['OS']).strip() if pd.notna(row['OS']) else ""
        codproduto = str(row['CODPRODUTO']).strip() if pd.notna(row['CODPRODUTO']) else ""
        
        # Criar chave para busca no dicionário
        chave = (nf_e, os_val, codproduto)
        
        # Buscar no dicionário
        if chave in qtde_ajustada_dict:
            peso_encontrado = qtde_ajustada_dict[chave]
            
            # VERIFICAR SE A COLUNA CF É "DEV" - SE FOR, TORNAR O VALOR NEGATIVO
            cf_valor = str(row['CF']).strip() if pd.notna(row['CF']) else ""
            if cf_valor == "DEV":
                peso_encontrado = -abs(peso_encontrado)  # Garante que seja negativo
                print(f"✅ Encontrado PESO para NF-E {nf_e}, OS {os_val}, CODPRODUTO {codproduto}: {peso_encontrado} (NEGATIVO - CF=DEV)")
            else:
                print(f"✅ Encontrado PESO para NF-E {nf_e}, OS {os_val}, CODPRODUTO {codproduto}: {peso_encontrado} (POSITIVO)")
            
            return peso_encontrado
        
        # Se não encontrou no dicionário, usar a lógica original como fallback
        if row['QTDE REAL'] <= 0:
            return row['QTDE REAL']
        
        data = row['DATA']
        
        if codproduto is None or data is None:
            return row['QTDE REAL']
            
        custo_info = custos_dict.get((codproduto, data), {})
        qtd = custo_info.get('QTD', 1)
        
        if qtd > 1:
            resultado = row['QTDE'] * qtd
        else:
            resultado = row['QTDE REAL']
            
        # Aplicar mesma condição de negativo para o fallback
        cf_valor = str(row['CF']).strip() if pd.notna(row['CF']) else ""
        if cf_valor == "DEV" and resultado > 0:
            resultado = -resultado
            
        return resultado
            
    except Exception as e:
        print(f"❌ Erro ao calcular QTDE AJUSTADA para linha: {e}")
        return row['QTDE REAL']

def calcular_qtde_real2(row):
    try:
        codproduto = str(row['CODPRODUTO']).strip() if pd.notna(row['CODPRODUTO']) else None
        data = row['DATA']
        
        if codproduto is None or data is None:
            return np.nan
            
        custo_info = custos_dict.get((codproduto, data), {})
        peso = custo_info.get('PESO', 1)
        
        if row['QTDE REAL'] < 0:
            return -row['QTDE AJUSTADA'] * peso
        else:
            return row['QTDE AJUSTADA'] * peso
    except:
        return np.nan

def buscar_custo(row):
    try:
        codproduto = str(row['CODPRODUTO']).strip() if pd.notna(row['CODPRODUTO']) else None
        data = row['DATA']
        
        if codproduto is None or data is None:
            return np.nan
            
        key = (codproduto, data)
        if key in custos_dict:
            custo = custos_dict[key].get('CUSTO', 0)
            return custo if custo != 0 else np.nan
        else:
            return np.nan
    except:
        return np.nan

def buscar_frete(row):
    if row['FANTASIA'] in ["PASSOS ALIMENTOS LTDA", "AGELLE ARMAZEM E LOGISTICA LTDA", 
                           "GEMEOS REORESENTACOES", "REAL DISTRIBUIDORA"]:
        return 0
        
    try:
        codproduto = str(row['CODPRODUTO']).strip() if pd.notna(row['CODPRODUTO']) else None
        data = row['DATA']
        
        if codproduto is None or data is None:
            return 0
            
        custo_info = custos_dict.get((codproduto, data), {})
        return custo_info.get('FRETE', 0)
    except:
        return 0

def buscar_producao(row):
    try:
        codproduto = str(row['CODPRODUTO']).strip() if pd.notna(row['CODPRODUTO']) else None
        data = row['DATA']
        
        if codproduto is None or data is None:
            return 0
            
        custo_info = custos_dict.get((codproduto, data), {})
        return custo_info.get('PRODUÇÃO', 0)
    except:
        return 0

# Aplicar funções
print("🔄 Calculando QTDE AJUSTADA...")
base_df['QTDE AJUSTADA'] = base_df.apply(calcular_qtde_ajustada, axis=1)
base_df['QTDE REAL2'] = base_df.apply(calcular_qtde_real2, axis=1)
base_df['CUSTO'] = base_df.apply(buscar_custo, axis=1)
base_df['Custo total'] = base_df['CUSTO'] * base_df['QTDE AJUSTADA']
base_df['Frete'] = base_df.apply(buscar_frete, axis=1)
base_df['Produção'] = base_df.apply(buscar_producao, axis=1)

# [O RESTO DO CÓDIGO PERMANECE EXATAMENTE IGUAL...]
# Escritório
if 'ESCRITORIO' in fechamento_sem_cancelados.columns:
    base_df['Escritório'] = fechamento_sem_cancelados['ESCRITORIO'].fillna(0) / 100
else:
    base_df['Escritório'] = 0

base_df['Escritório'] = base_df['Escritório'].apply(lambda x: 0.035 if abs(x - 0.04) < 0.001 else x)

# Desconto
base_df['Desc Perc'] = 0
if 'DESCONTO' in fechamento_sem_cancelados.columns:
    for i, row in fechamento_sem_cancelados.iterrows():
        if i < len(base_df):
            desconto_val = row['DESCONTO']
            if pd.notna(desconto_val) and str(desconto_val).strip() != '':
                try:
                    base_df.at[i, 'Desc Perc'] = float(str(desconto_val).replace(',', '.').strip()) / 100
                except:
                    base_df.at[i, 'Desc Perc'] = 0

base_df['Desc. Valor'] = base_df.apply(
    lambda row: 0 if (row['CF'] == "DEV" or row['GRUPO'] == "TENDA") 
    else row['QTDE AJUSTADA'] * row['Preço Venda'] * row['Desc Perc'], axis=1
)

# Fat. Bruto
base_df['Fat. Bruto'] = base_df.apply(
    lambda row: -row['QTDE AJUSTADA'] * row['Preço Venda'] if row['CF'] == "DEV"
    else row['QTDE AJUSTADA'] * row['Preço Venda'], axis=1
)

# Aliq Icms
base_df['Aliq Icms'] = base_df.apply(
    lambda row: round(row['VL ICMS'] / row['Fat. Bruto'], 2) if (row['Fat. Bruto'] != 0 and pd.notna(row['VL ICMS'])) 
    else 0, axis=1
)
base_df['Aliq Icms'] = base_df['Aliq Icms'].replace([np.inf, -np.inf], 0)

# Custo real
base_df['Custo real'] = base_df.apply(
    lambda row: 0 if (pd.isna(row['QTDE AJUSTADA']) or row['QTDE AJUSTADA'] <= 0 or 
                     pd.isna(row['CUSTO']) or pd.isna(row['Aliq Icms']))
    else round(row['CUSTO'] - (row['CUSTO'] * row['Aliq Icms']), 2), axis=1
)

# Fat Liquido
base_df['Fat Liquido'] = base_df.apply(
    lambda row: row['QTDE AJUSTADA'] * (row['Preço Venda'] - row['Preço Venda'] * row['Desc Perc']) if row['CF'] != "DEV"
    else -row['QTDE AJUSTADA'] * (row['Preço Venda'] - row['Preço Venda'] * row['Desc Perc']), axis=1
)

# Aniversário
base_df['Aniversário'] = base_df.apply(
    lambda row: 0 if row['CF'] == "DEV" else row['Fat. Bruto'] * 0.01, axis=1
)

# Comissão Kg
def calcular_comissao_kg(row):
    try:
        if row['CF'] == "DEV":
            comissao = -(row['Preço Venda'] * row['P. Com'])
            return comissao
        
        vendedor = str(row['VENDEDOR']).strip() if pd.notna(row['VENDEDOR']) else ''
        codproduto = str(row['CODPRODUTO']).strip() if pd.notna(row['CODPRODUTO']) else ''
        grupo = str(row['GRUPO']).strip() if pd.notna(row['GRUPO']) else ''
        
        # Regras específicas por vendedor
        regras_vendedores = {
            "LUIZ FERNANDO VOLTERO BARBOSA": {"812": {"CHAMA": 3, "PARANA": 3, "REDE PLUS": 2}},
            "FELIPE RAMALHO GOMES": {"700": {"PEDREIRA": 2, "BERGAMINI": 0.5}},
            "VALDENIR VOLTERO": {"812": {"RICOY": 1}, "937": {"RICOY": 0.5}, "1624": {"RICOY": 0.5}},
            "ROSE VOLTERO": {"812": 2},
            "VERA LUCIA MUNIZ": {"812": 2},
            "PAMELA FERREIRA VIEIRA": {"812": 2}
        }
        
        if vendedor in regras_vendedores:
            if codproduto in regras_vendedores[vendedor]:
                regra = regras_vendedores[vendedor][codproduto]
                if isinstance(regra, dict):
                    return regra.get(grupo, row['Preço Venda'] * row['P. Com'])
                else:
                    return regra
        
        if row['GRUPO'] != "LOURENCINI" or lourencini.empty:
            return (row['Preço Venda'] * row['P. Com'])
        
        # Lógica LOURENCINI
        data_venda = row['DATA']
        preco_venda = row['Preço Venda']
        codproduto = str(row['CODPRODUTO']).strip()
        
        if not codproduto or codproduto == 'nan' or pd.isna(data_venda) or pd.isna(preco_venda) or preco_venda == 0:
            return (row['Preço Venda'] * row['P. Com'])
        
        lourencini_filtrado = lourencini[lourencini['COD'] == codproduto]
        if lourencini_filtrado.empty:
            return (row['Preço Venda'] * row['P. Com'])
        
        data_venda_dt = pd.Timestamp(data_venda)
        lourencini_row = None
        
        if 'Data_fim' in lourencini_filtrado.columns:
            lourencini_periodo = lourencini_filtrado[
                (lourencini_filtrado['Data'] <= data_venda_dt) & 
                (lourencini_filtrado['Data_fim'] >= data_venda_dt)
            ]
            if not lourencini_periodo.empty:
                lourencini_row = lourencini_periodo.iloc[0]
        
        if lourencini_row is None:
            lourencini_anteriores = lourencini_filtrado[lourencini_filtrado['Data'] <= data_venda_dt]
            if not lourencini_anteriores.empty:
                lourencini_anteriores = lourencini_anteriores.sort_values('Data', ascending=False)
                lourencini_row = lourencini_anteriores.iloc[0]
            else:
                lourencini_filtrado = lourencini_filtrado.sort_values('Data', ascending=True)
                lourencini_row = lourencini_filtrado.iloc[0]
        
        colunas_comissao = ['0,2', '0,3', '0,5', '0,7', '1']
        for coluna in colunas_comissao:
            valor_na_tabela = lourencini_row[coluna]
            if pd.notna(valor_na_tabela) and abs(preco_venda - valor_na_tabela) < 0.01:
                return float(coluna.replace(',', '.'))
        
        return (row['Preço Venda'] * row['P. Com'])
        
    except Exception:
        return (row['Preço Venda'] * row['P. Com'])

base_df['P. Com'] = 0
base_df['Comissão Kg'] = base_df.apply(calcular_comissao_kg, axis=1)

# Regras de comissão fixa
def criar_regras_comissao_fixa():
    return {
        'geral': {
            0.00: {
                'grupos': ['AKKI', 'ANDORINHA', 'BERGAMINI', 'DA PRACA', 'DOVALE', 'MERCADAO', 'REIMBERG', 'SEMAR', 'TRIMAIS', 'VOVO ZUZU', 'BENGALA', 'OURINHOS'],
                'razoes': ['COMERCIO DE CARNES E ROTISSERIE DUTRA LTDA', 'DISTRIBUIDORA E COMERCIO UAI SP LTDA', "GARFETO'S FORNECIMENTO DE REFEICOES LTDA", "LATICINIO SOBERANO LTDA VILA ALPINA", "SAO LORENZO ALIMENTOS LTDA", "QUE DELICIA MENDES COMERCIO DE ALIMENTOS", "MARIANA OLIVEIRA MAZZEI", "LS SANTOS COMERCIO DE ALIMENTOS LTDA"]
            },
            0.03: {'grupos': ['CALVO', 'CHAMA', 'ESTRELA AZUL', 'TENDA', 'HIGAS']},
            0.01: {'razoes': ['SHOPPING FARTURA VALINHOS COMERCIO LTDA']}
        }
    }

def aplicar_regras_comissao_fixa(row):
    try:
        regras = criar_regras_comissao_fixa()
        grupo = str(row['GRUPO']).strip() if pd.notna(row['GRUPO']) else ''
        razao = str(row['RAZAO']).strip() if pd.notna(row['RAZAO']) else ''
        fantasia = str(row['FANTASIA']).strip() if pd.notna(row['FANTASIA']) else ''
        
        for comissao, regra in regras['geral'].items():
            if 'grupos' in regra and grupo in regra['grupos']:
                return comissao
            if 'razoes' in regra and (razao in regra['razoes'] or fantasia in regra['razoes']):
                return comissao
        return None
    except:
        return None

def aplicar_ofertas_comissao(row):
    try:
        # Verificar se temos os dataframes de ofertas carregados
        if ofertas_vog.empty:
            return 0.03  # Default se não tiver ofertas carregadas
        
        codproduto = str(row['CODPRODUTO']).strip() if pd.notna(row['CODPRODUTO']) else ''
        data_venda = row['DATA']
        preco_venda = row['Preço Venda']
        
        if not codproduto or codproduto == 'nan' or pd.isna(data_venda) or pd.isna(preco_venda) or preco_venda == 0:
            return 0.03
        
        # Converter data_venda para Timestamp se necessário
        if isinstance(data_venda, date):
            data_venda = pd.Timestamp(data_venda)
        
        # Converter codproduto para inteiro para comparação
        try:
            codproduto_int = int(float(codproduto))
        except (ValueError, TypeError):
            return 0.03
        
        # ESTRATÉGIA 1: Buscar na aba OFF_VOG (com DT_REF_OFF)
        if 'COD' in ofertas_vog.columns and 'DT_REF_OFF' in ofertas_vog.columns:
            # Filtrar ofertas para o código do produto
            ofertas_cod = ofertas_vog[ofertas_vog['COD'] == codproduto_int].copy()
            
            if not ofertas_cod.empty:
                
                # Converter e ordenar datas
                ofertas_cod['DT_REF_OFF'] = pd.to_datetime(ofertas_cod['DT_REF_OFF'], errors='coerce')
                ofertas_cod = ofertas_cod.dropna(subset=['DT_REF_OFF'])
                ofertas_cod = ofertas_cod.sort_values('DT_REF_OFF', ascending=False)
                
                # Encontrar oferta mais próxima (mais recente anterior ou igual à data da venda)
                ofertas_validas = ofertas_cod[ofertas_cod['DT_REF_OFF'] <= data_venda]
                
                if not ofertas_validas.empty:
                    oferta_mais_recente = ofertas_validas.iloc[0]  # Já está ordenado
                    
                    # Verificar colunas de preço disponíveis - ORDEM DE PRIORIDADE: 3% -> 1%
                    if '3%' in oferta_mais_recente and pd.notna(oferta_mais_recente['3%']):
                        preco_oferta = oferta_mais_recente['3%']
                        if preco_venda >= preco_oferta:
                            return 0.03
                    
                    # Se não atendeu ao 3%, verificar 1%
                    if '1%' in oferta_mais_recente and pd.notna(oferta_mais_recente['1%']):
                        preco_oferta = oferta_mais_recente['1%']
                        if preco_venda >= preco_oferta:
                            return 0.01
        
        # ESTRATÉGIA 2: Buscar na aba OFF_VOG_CB (com DT_REF_OFF_CB)
        if 'CD_PROD' in ofertas_vog.columns and 'DT_REF_OFF_CB' in ofertas_vog.columns:
            # Filtrar ofertas para o código do produto
            ofertas_cod = ofertas_vog[ofertas_vog['CD_PROD'] == codproduto_int].copy()
            
            if not ofertas_cod.empty:
                
                # Converter e ordenar datas
                ofertas_cod['DT_REF_OFF_CB'] = pd.to_datetime(ofertas_cod['DT_REF_OFF_CB'], errors='coerce')
                ofertas_cod = ofertas_cod.dropna(subset=['DT_REF_OFF_CB'])
                ofertas_cod = ofertas_cod.sort_values('DT_REF_OFF_CB', ascending=False)
                
                # Encontrar oferta mais próxima
                ofertas_validas = ofertas_cod[ofertas_cod['DT_REF_OFF_CB'] <= data_venda]
                
                if not ofertas_validas.empty:
                    oferta_mais_recente = ofertas_validas.iloc[0]
                    
                    # Verificar colunas de preço disponíveis - ORDEM DE PRIORIDADE: 2% -> 1%
                    if '2%' in oferta_mais_recente and pd.notna(oferta_mais_recente['2%']):
                        preco_oferta = oferta_mais_recente['2%']
                        if preco_venda >= preco_oferta:
                            return 0.02
                    
                    # Se não atendeu ao 2%, verificar 1%
                    if '1%' in oferta_mais_recente and pd.notna(oferta_mais_recente['1%']):
                        preco_oferta = oferta_mais_recente['1%']
                        if preco_venda >= preco_oferta:
                            return 0.01
        
        # Se não encontrou em nenhuma tabela de ofertas, aplicar regra padrão por grupo
        grupo = str(row['GRUPO']).strip() if pd.notna(row['GRUPO']) else ''
        comissao_padrao = 0.02 if grupo == "CORTES BOVINOS" else 0.03
        return comissao_padrao
        
    except Exception as e:
        print(f"❌ Erro ao aplicar ofertas comissão para código {codproduto}: {e}")
        # Fallback para regra padrão por grupo
        grupo = str(row['GRUPO']).strip() if pd.notna(row['GRUPO']) else ''
        return 0.02 if grupo == "CORTES BOVINOS" else 0.03

def calcular_p_com_com_regras_fixas(row):
    try:
        comissao_kg = row['Comissão Kg']
        preco_venda = row['Preço Venda']
        
        if comissao_kg > 0 and preco_venda > 0:
            return comissao_kg / preco_venda
        elif comissao_kg <= 0:
            comissao_regras_fixas = aplicar_regras_comissao_fixa(row)
            if comissao_regras_fixas is not None:
                return comissao_regras_fixas
            else:
                return aplicar_ofertas_comissao(row)
        else:
            return 0
    except:
        return aplicar_ofertas_comissao(row)

base_df['P. Com'] = base_df.apply(calcular_p_com_com_regras_fixas, axis=1)

# Colunas restantes
base_df['Comissão Real'] = base_df.apply(
    lambda row: row['Comissão Kg'] * row['QTDE AJUSTADA'] if row['Preço Venda'] > 0 
    else -(row['Comissão Kg'] * row['QTDE AJUSTADA']), axis=1
)

base_df['Coleta Esc'] = base_df['Fat. Bruto'] * base_df['Escritório']
base_df['Frete Real'] = base_df['QTDE REAL2'] * base_df['Frete']

base_df['comissão'] = base_df.apply(
    lambda row: row['Comissão Kg'] / row['Preço Venda'] if row['Preço Venda'] > 0
    else -row['Comissão Kg'] / row['Preço Venda'] if row['Preço Venda'] < 0
    else 0, axis=1
)

base_df['Escr.'] = base_df.apply(
    lambda row: row['Coleta Esc'] / row['Fat. Bruto'] if row['Fat. Bruto'] != 0
    else 0, axis=1
)

base_df['frete'] = base_df.apply(
    lambda row: row['Frete Real'] / row['Fat. Bruto'] if row['Fat. Bruto'] != 0
    else 0, axis=1
)

base_df['TP'] = base_df.apply(
    lambda row: "BIG BACON" if str(row['CODPRODUTO']) == "700"
    else "REFFINATO" if row['GRUPO PRODUTO'] in ["SALGADOS SUINOS A GRANEL", "SALGADOS SUINOS EMBALADOS"]
    else "MIX", axis=1
)

base_df['CANC'] = base_df['NF-E'].apply(lambda x: "SIM" if x in notas_canceladas else "")

base_df['Armazenagem'] = base_df.apply(
    lambda row: (row['Fat. Bruto'] * row['P. Com']) / row['QTDE AJUSTADA'] if row['QTDE AJUSTADA'] != 0
    else 0, axis=1
)

def buscar_comissao_regra(row):
    try:
        if pd.notna(row['OS']) and pd.notna(row['NF-E']) and pd.notna(row['CODPRODUTO']):
            key = (int(row['OS']), int(row['NF-E']), int(row['CODPRODUTO']))
            return comissao_regra_dict.get(key, 0)
        else:
            return 0
    except:
        return 0

base_df['Comissão por Regra'] = base_df.apply(buscar_comissao_regra, axis=1)
base_df['Coluna2'] = base_df['Comissão por Regra'] == base_df['Comissão Kg']
base_df['FRETE - LUC/PREJ'] = base_df['QTDE AJUSTADA'] * base_df['Frete']

def buscar_desc_fec(row):
    try:
        pk = str(row['OS']) + "_" + str(row['NF-E']) + "_" + str(row['CODPRODUTO'])
        return fechamento_pk_dict.get(pk, {}).get('DESCONTO_VERIFICADO', np.nan)
    except:
        return np.nan

base_df['DESC FEC'] = base_df.apply(buscar_desc_fec, axis=1)

def buscar_esc_fec(row):
    try:
        pk = str(row['OS']) + "_" + str(row['NF-E']) + "_" + str(row['CODPRODUTO'])
        esc_fec_value = fechamento_pk_dict.get(pk, {}).get('ESCRITORIO', np.nan)
        if pd.isna(esc_fec_value):
            return np.nan
        esc_fec_value = float(esc_fec_value)
        if abs(esc_fec_value - 4.0) < 0.001:
            esc_fec_value = 3.5
        return esc_fec_value / 100
    except:
        return np.nan

base_df['ESC FEC'] = base_df.apply(buscar_esc_fec, axis=1)

def buscar_icms_fec(row):
    try:
        pk = str(row['OS']) + "_" + str(row['NF-E']) + "_" + str(row['CODPRODUTO'])
        return fechamento_pk_dict.get(pk, {}).get('VLR ICMS', np.nan)
    except:
        return np.nan

base_df['ICMS FEC'] = base_df.apply(buscar_icms_fec, axis=1)

def buscar_prc_vend_fev(row):
    try:
        pk = str(row['OS']) + "_" + str(row['NF-E']) + "_" + str(row['CODPRODUTO'])
        return fechamento_pk_dict.get(pk, {}).get('PRECO VENDA', np.nan)
    except:
        return np.nan

base_df['PRC VEND FEV'] = base_df.apply(buscar_prc_vend_fev, axis=1)

base_df['DESC'] = base_df.apply(
    lambda row: (row['DESC FEC'] / 100) == row['Desc Perc'] if pd.notna(row['DESC FEC']) else False, axis=1
)

base_df['ESC'] = base_df.apply(
    lambda row: (row['ESC FEC'] / 100) == row['Escritório'] if pd.notna(row['ESC FEC']) else False, axis=1
)

base_df['ICMS'] = base_df.apply(
    lambda row: row['ICMS FEC'] == row['VL ICMS'] if pd.notna(row['ICMS FEC']) else False, axis=1
)

base_df['PRC VEND'] = base_df.apply(
    lambda row: row['PRC VEND FEV'] == row['Preço Venda'] if pd.notna(row['PRC VEND FEV']) else False, axis=1
)

def buscar_descricao_1(row):
    try:
        nf_e = int(row['NF-E']) if pd.notna(row['NF-E']) else None
        if nf_e is not None and nf_e in fechamento_nf_dict:
            return fechamento_nf_dict[nf_e]
        else:
            return ""
    except:
        return ""

base_df['DESCRIÇÃO_1'] = base_df.apply(buscar_descricao_1, axis=1)

base_df['MOV ENC'] = base_df.apply(
    lambda row: "ENCONTRADO" if any([str(row['OS']) == str(venda[0]) and str(row['NF-E']) == str(venda[1]) for venda in vendas_var])
    else "NÃO ENCONTRADO", axis=1
)

base_df['CUST + IMP'] = base_df['Custo real'] * base_df['QTDE AJUSTADA']
base_df['CUST PROD'] = base_df['QTDE AJUSTADA'] * base_df['Produção']
base_df['COM BRUTA'] = base_df['QTDE AJUSTADA'] * base_df['P. Com'] * base_df['Preço Venda']
base_df['Coluna1'] = (round(base_df['COM BRUTA'], 2) == round(base_df['Comissão Real'], 2))

base_df['Custo divergente'] = base_df.apply(
    lambda row: "CORRETO" if (row['QTDE'] > 0 and row['CUSTO EM SISTEMA'] == row['CUSTO']) else "DIVERGENTE", axis=1
)

base_df['Lucro / Prej.'] = base_df.apply(
    lambda row: row['Fat. Bruto'] - (
        (row['QTDE AJUSTADA'] * row['Custo real']) + 
        row['Val Pis'] + 
        row['VLRCOFINS'] + 
        row['IRPJ'] + 
        row['CSLL'] + 
        row['VL ICMS'] + 
        row['Desc. Valor'] + 
        (row['QTDE AJUSTADA'] * row['Frete']) + 
        (row['QTDE AJUSTADA'] * row['Produção']) + 
        (row['P. Com'] * row['Preço Venda'] * row['QTDE AJUSTADA']) + 
        row['Aniversário']
    ) - (row['Fat. Bruto'] * row['Escritório']), 
    axis=1
)

base_df['Margem'] = base_df.apply(
    lambda row: (row['Lucro / Prej.'] / row['Fat Liquido']) if row['Fat Liquido'] != 0 else 0, axis=1
)

base_df['INCL.'] = ""

def buscar_descricao_2(row):
    try:
        pk = str(row['OS']) + "_" + str(row['NF-E']) + "_" + str(row['CODPRODUTO'])
        return fechamento_pk_dict.get(pk, {}).get('MOV_V2', "")
    except:
        return ""

base_df['DESCRIÇÃO_2'] = base_df.apply(buscar_descricao_2, axis=1)

# Reordenar colunas
colunas_ordenadas = [
    'CF', 'RAZAO', 'FANTASIA', 'GRUPO', 'OS', 'NF-E', 'CF_NF', 'DATA', 'VENDEDOR',
    'CODPRODUTO', 'GRUPO PRODUTO', 'DESCRICAO', 'QTDE', 'QTDE REAL', 'CUSTO EM SISTEMA',
    'QTDE AJUSTADA', 'QTDE REAL2', 'CUSTO', 'Custo real', 'Custo total', 'Frete', 'Produção',  # "Custo total" ADICIONADA AQUI
    'Escritório', 'Comissão Kg', 'P. Com', 'Aniversário', 'Val Pis', 'VLRCOFINS',
    'IRPJ', 'CSLL', 'VL ICMS', 'Aliq Icms', 'Desc Perc', 'Desc. Valor', 'Preço Venda',
    'Fat Liquido', 'Fat. Bruto', 'Lucro / Prej.', 'Margem', 'Quinzena', 'Comissão Real',
    'Coleta Esc', 'Frete Real', 'INCL.', 'comissão', 'Escr.', 'frete', 'Custo divergente',
    'TP', 'CANC', 'Armazenagem', 'Comissão por Regra', 'PK', 'Coluna2', 'FRETE - LUC/PREJ',
    'CUST + IMP', 'CUST PROD', 'COM BRUTA', 'DESC FEC', 'ESC FEC', 'ICMS FEC', 'PRC VEND FEV',
    'DESC', 'ESC', 'ICMS', 'PRC VEND', 'Coluna1', 'DESCRIÇÃO_1', 'DESCRIÇÃO_2'
]

colunas_existentes = [col for col in colunas_ordenadas if col in base_df.columns]
base_df = base_df[colunas_existentes]
base_df = base_df.fillna("")


# Salvar arquivos
print("💾 Salvando arquivos...")
output_path = f"C:\\Users\\win11\\Downloads\\Margem_{data_nome_arquivo}.xlsx"

with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    # Salvar cada planilha com formatação de fonte
    base_df.to_excel(writer, sheet_name='base (3,5%)', index=False)
    ofertas_vog.to_excel(writer, sheet_name='OFERTAS_VOG', index=False)
    custos_produtos.to_excel(writer, sheet_name='PRECOS', index=False)
    cancelados.to_excel(writer, sheet_name='Base_cancelamento', index=False)
    devolucoes.to_excel(writer, sheet_name='Base_movimentacao', index=False)
    fechamento.to_excel(writer, sheet_name='Base_Fechamento', index=False)
    
    if not lourencini.empty:
        lourencini.to_excel(writer, sheet_name='LOURENCINI', index=False)
    
    # Aplicar formatação de fonte tamanho 10 para todas as planilhas
    workbook = writer.book
    font_size = 10
    
    # Definir colunas para centralizar (INCLUINDO as de porcentagem)
    colunas_para_centralizar = [
        'CF', 'OS', 'NF-E', 'CF_NF', 'DATA', 'CODPRODUTO', 'QTDE', 'QTDE REAL', 
        'CUSTO EM SISTEMA', 'QTDE AJUSTADA', 'QTDE REAL2', 'Escritório', 'P. Com', 
        'Desc Perc', 'Margem', 'Quinzena', 'INCL.', 'comissão', 'Escr.', 'frete', 
        'Custo divergente', 'TP', 'PK', 'Coluna2', 'FRETE - LUC/PREJ', 'CUST PROD', 
        'COM BRUTA', 'DESC FEC', 'ESC FEC', 'ICMS FEC', 'PRC VEND FEV', 'DESC', 
        'ESC', 'ICMS', 'PRC VEND', 'Coluna1', 'DESCRIÇÃO_1', 'DESCRIÇÃO_2'
    ]
    
    # Definir colunas para formatação monetária (formato Real brasileiro)
    colunas_monetarias = [
        'CUSTO', 'Custo real', 'Custo total', 'Frete', 'Produção', 'Comissão Kg', 'Aniversário',  # "Custo total" ADICIONADA AQUI
        'VL ICMS', 'Desc. Valor', 'Preço Venda', 'Fat Liquido', 'Fat. Bruto',
        'Lucro / Prej.', 'Comissão Real', 'Coleta Esc', 'Frete Real',
        'Armazenagem', 'Comissão por Regra', 'CUST + IMP'
    ]
    
    # Definir colunas para formatação de porcentagem com 2 casas decimais
    colunas_porcentagem = [
        'Escritório', 'P. Com', 'Desc Perc', 'Margem', 'comissão', 'Escr.', 'frete',
        'DESC FEC', 'ESC FEC'
    ]
    
    for sheet_name in writer.sheets:
        worksheet = writer.sheets[sheet_name]
        
        # Primeiro, encontrar os índices das colunas
        if sheet_name == 'base (3,5%)':
            col_indices_monetarias = {}
            col_indices_porcentagem = {}
            col_indices_centralizar = {}
            
            for col_num in range(1, worksheet.max_column + 1):
                col_name = worksheet.cell(row=1, column=col_num).value
                if col_name in colunas_monetarias:
                    col_indices_monetarias[col_num] = col_name
                if col_name in colunas_porcentagem:
                    col_indices_porcentagem[col_num] = col_name
                if col_name in colunas_para_centralizar:
                    col_indices_centralizar[col_num] = col_name
        
        # APLICAR CENTRALIZAÇÃO PARA AS COLUNAS ESPECIFICADAS (incluindo porcentagem)
        if sheet_name == 'base (3,5%)':
            for col_num in col_indices_centralizar:
                col_letter = openpyxl.utils.get_column_letter(col_num)
                
                # Aplicar alinhamento centralizado para TODA a coluna (incluindo cabeçalho)
                for row_num in range(1, worksheet.max_row + 1):
                    cell = worksheet[f'{col_letter}{row_num}']
                    cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # Aplicar formatação monetária para TODAS as células das colunas monetárias
        if sheet_name == 'base (3,5%)':
            for col_num in col_indices_monetarias:
                col_letter = openpyxl.utils.get_column_letter(col_num)
                
                # Aplicar o formato de moeda brasileiro completo para toda a coluna
                for row_num in range(2, worksheet.max_row + 1):
                    cell = worksheet[f'{col_letter}{row_num}']
                    if cell.value is not None:
                        try:
                            float(cell.value)
                            # Formato com R$ à esquerda e número à direita
                            cell.number_format = '"R$"* #,##0.00;[Red]"R$"* -#,##0.00;"R$"* -'
                            # Manter a centralização se a coluna também estiver na lista de centralizar
                            if col_num in col_indices_centralizar:
                                cell.alignment = Alignment(horizontal='center', vertical='center', 
                                                         number_format='"R$"* #,##0.00;[Red]"R$"* -#,##0.00;"R$"* -')
                        except (ValueError, TypeError):
                            pass
        
        # Aplicar formatação de porcentagem com 2 casas decimais E centralização
        if sheet_name == 'base (3,5%)':
            for col_num in col_indices_porcentagem:
                col_letter = openpyxl.utils.get_column_letter(col_num)
                
                # Aplicar o formato de porcentagem para toda a coluna COM centralização
                for row_num in range(2, worksheet.max_row + 1):
                    cell = worksheet[f'{col_letter}{row_num}']
                    if cell.value is not None:
                        try:
                            float(cell.value)
                            # Formato de porcentagem com 2 casas decimais E centralização
                            cell.number_format = '0.00%'
                            # Garantir que as colunas de porcentagem também fiquem centralizadas
                            cell.alignment = Alignment(horizontal='center', vertical='center', 
                                                     number_format='0.00%')
                        except (ValueError, TypeError):
                            pass
        
        # Aplicar fonte tamanho 10 para todas as células
        for row in worksheet.iter_rows():
            for cell in row:
                cell.font = openpyxl.styles.Font(size=font_size)
                # Garantir que o cabeçalho das colunas de porcentagem também fique centralizado
                if cell.row == 1 and cell.column in col_indices_porcentagem:
                    cell.alignment = Alignment(horizontal='center', vertical='center')
        
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter

            for cell in column:
                try:
                    if cell.value is not None:
                        # Considerar o comprimento do conteúdo, mas com limites razoáveis
                        cell_length = len(str(cell.value))

                        # Aplicar limites diferentes baseados no tipo de conteúdo estimado
                        if any(char.isdigit() for char in str(cell.value)) and not any(char.isalpha() for char in str(cell.value)):
                            # Provavelmente é número/data - largura menor
                            max_length = max(min(cell_length, 12), max_length)
                        else:
                            # É texto - largura maior mas ainda limitada
                            max_length = max(min(cell_length, 25), max_length)
                except:
                    pass
                
            # Ajustar largura final com limites sensatos
            adjusted_width = min(max_length + 2, 30)  # Máximo de 30
            adjusted_width = max(adjusted_width, 10)   # Mínimo de 10 para legibilidade

            worksheet.column_dimensions[column_letter].width = adjusted_width


# Salvar JSON (código original mantido)
def default_serializer(obj):
    if isinstance(obj, (np.integer, int)):
        return int(obj)
    elif isinstance(obj, (np.floating, float)):
        return float(obj)
    elif isinstance(obj, (np.ndarray, pd.Series)):
        return obj.tolist()
    elif isinstance(obj, pd.Timestamp):
        return obj.isoformat() if not pd.isna(obj) else ""
    elif isinstance(obj, date):
        return obj.isoformat() if obj is not None else ""
    elif pd.isna(obj):
        return None
    elif obj in [np.inf, -np.inf]:
        return None
    elif callable(obj):
        return ""
    elif hasattr(obj, '__call__'):
        return ""
    raise TypeError(f"Type {type(obj)} not serializable")

json_path = f"C:\\Users\\win11\\Downloads\\Margem_{data_nome_arquivo}.json"
base_df_clean = base_df.copy()

with open(json_path, 'w', encoding='utf-8') as f:
    json.dump(base_df_clean.to_dict(orient='records'), f, ensure_ascii=False, indent=4, default=default_serializer)

print("✅ PROCESSAMENTO CONCLUÍDO!")
print(f"📄 Arquivo Excel: {output_path}")
print(f"📄 Arquivo JSON: {json_path}")
print(f"📊 Total de registros processados: {len(base_df)}")