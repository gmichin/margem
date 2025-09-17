import pandas as pd
import numpy as np
from datetime import date
import warnings
import json

warnings.filterwarnings('ignore')

# Data fixa para o nome do arquivo
data_nome_arquivo = "050825"

# Fechamento.csv
fechamento = pd.read_csv(r"C:\Users\win11\Downloads\fechamento.csv", sep=';', encoding='utf-8', decimal=',', thousands='.')

# Cancelados.csv (pula as 2 primeiras linhas)
cancelados = pd.read_csv(r"C:\Users\win11\Downloads\cancelados.csv", sep=';', encoding='utf-8', decimal=',', thousands='.', skiprows=2)

# Devoluções.csv
devolucoes = pd.read_csv(r"C:\Users\win11\Downloads\movimentação.csv", sep=';', encoding='utf-8', decimal=',', thousands='.')

# Custos de produtos - Julho.xlsx
custos_produtos = pd.read_excel(r"C:\Users\win11\Downloads\Custos de produtos - Julho.xlsx", 
                               sheet_name='Base',
                               dtype=str)  # Ler tudo como string primeiro

# RENOMEAR COLUNAS CORRETAMENTE
rename_mapping = {
    'PRODUTO': 'CODPRODUTO',
    'DATA': 'DATA',
    'PCS': 'QTDE',
    'KGS': 'PESO_KGS',
    'CUSTO': 'CUSTO',
    'FRETE': 'FRETE',
    'PRODUÇÃO': 'PRODUÇÃO',
    'TOTAL': 'TOTAL',
    'QTD': 'QTD',
    'PESO': 'PESO'
}

custos_produtos = custos_produtos.rename(columns=rename_mapping)

# CONVERTER COLUNAS NUMÉRICAS - MÉTODO MAIS ROBUSTO
colunas_numericas = ['PESO_KGS', 'CUSTO', 'FRETE', 'PRODUÇÃO', 'TOTAL', 'QTD', 'PESO']

for coluna in colunas_numericas:
    if coluna in custos_produtos.columns:
        try:
            # Converter para string primeiro se não for já
            if custos_produtos[coluna].dtype != 'object':  # CORREÇÃO: Acessar a coluna específica
                custos_produtos[coluna] = custos_produtos[coluna].astype(str)
            
            # Substituir vírgula por ponto e remover possíveis espaços
            custos_produtos[coluna] = custos_produtos[coluna].apply(
                lambda x: str(x).replace(',', '.').replace(' ', '') if pd.notna(x) else x
            )
            
            # Converter para numérico
            custos_produtos[coluna] = pd.to_numeric(custos_produtos[coluna], errors='coerce')
            
        except Exception as e:
            print(f"Erro ao converter coluna {coluna}: {e}")
    else:
        print(f"Coluna {coluna} não encontrada no DataFrame")

# Converter DATA para datetime
custos_produtos['DATA'] = pd.to_datetime(custos_produtos['DATA'], errors='coerce', dayfirst=True)

# Converter CODPRODUTO para string
custos_produtos['CODPRODUTO'] = custos_produtos['CODPRODUTO'].astype(str).str.strip()

# OFERTAS_VOG.xlsx
ofertas_vog = pd.read_excel(r"C:\Users\win11\Downloads\OFERTAS_VOG.xlsx")

# Converter colunas numéricas do fechamento
numeric_columns_fechamento = ['ROMANEIO', 'NF-E', 'CF_NF', 'CODPRODUTO', 'QTDE', 'QTDE REAL', 'CUSTO', 
                             'FRETE', 'PRODUCAO', 'ESCRITORIO', 'P.COM', 'ANIVERSARIO', 'VLR PIS', 
                             'VLR COFINS', 'IRPJ', 'CSLL', 'VLR ICMS', 'ALIQ ICMS', 'DESCONTO', 
                             'VLR DESCONTO', 'PRECO VENDA', 'FAT LIQUIDO', 'FAT BRUTO', 'LUCRO', 'MARGEM', 
                             'QTD POR EMB', 'FATOR DE CONVERSAO']

# 1. Notas canceladas
notas_canceladas = cancelados['NUMERO'].tolist()

# 2. Devoluções (DESCRICAO = "DEV VENDA C/ FIN S/ EST" ou HISTORICO = "68")
devolucoes_filtro = devolucoes[
    (devolucoes['DESCRICAO'] == "DEV VENDA C/ FIN S/ EST") | 
    (devolucoes['HISTORICO'] == "68")
]
devolucoes_var = devolucoes_filtro[['ROMANEIO', 'NOTA FISCAL']].values.tolist()

# 3. Vendas (DESCRICAO = "VENDA" ou HISTORICO = "51")
vendas_filtro = devolucoes[
    (devolucoes['DESCRICAO'] == "VENDA") | 
    (devolucoes['HISTORICO'] == "51")
]
vendas_var = vendas_filtro[['ROMANEIO', 'NOTA FISCAL']].values.tolist()

# Filtrar fechamento removendo notas canceladas
fechamento_sem_cancelados = fechamento[~fechamento['NF-E'].isin(notas_canceladas)].copy()

# Criar dicionário de custos - VERSÃO CORRIGIDA
custos_dict = {}
custos_produtos_sem_data = 0
custos_produtos_sem_codigo = 0

# Converter para lista de dicionários para evitar problemas com iterrows()
custos_data = custos_produtos.to_dict('records')

for i, row in enumerate(custos_data):
    try:
        # Verificar DATA - abordagem mais segura
        data_value = row.get('DATA', None)
        if data_value is None or pd.isna(data_value) or (isinstance(data_value, str) and data_value.strip() == ''):
            custos_produtos_sem_data += 1
            continue
            
        # Verificar CODPRODUTO - abordagem mais segura
        codproduto_value = row.get('CODPRODUTO', None)
        if codproduto_value is None or pd.isna(codproduto_value) or (isinstance(codproduto_value, str) and codproduto_value.strip() == ''):
            custos_produtos_sem_codigo += 1
            continue
        
        # Converter para os tipos corretos
        try:
            codproduto = str(codproduto_value).strip()
            data_key = pd.to_datetime(data_value, errors='coerce', dayfirst=True)
            if pd.isna(data_key):
                continue
            data_key = data_key.date()
            
            # Converter valores numéricos com tratamento de erro
            custo_val = float(row.get('CUSTO', 0)) if pd.notna(row.get('CUSTO')) and str(row.get('CUSTO', '')).strip() != '' else 0
            peso_val = float(row.get('PESO', 1)) if pd.notna(row.get('PESO')) and str(row.get('PESO', '')).strip() != '' else 1
            producao_val = float(row.get('PRODUÇÃO', 0)) if pd.notna(row.get('PRODUÇÃO')) and str(row.get('PRODUÇÃO', '')).strip() != '' else 0
            frete_val = float(row.get('FRETE', 0)) if pd.notna(row.get('FRETE')) and str(row.get('FRETE', '')).strip() != '' else 0
            qtd_val = float(row.get('QTD', 0)) if pd.notna(row.get('QTD')) and str(row.get('QTD', '')).strip() != '' else 0
            
            custos_dict[(codproduto, data_key)] = {
                'QTD': qtd_val,
                'PESO': peso_val,
                'CUSTO': custo_val,
                'FRETE': frete_val,
                'PRODUÇÃO': producao_val
            }
            
        except Exception as conv_error:
            print(f"Erro de conversão na linha {i}: {conv_error}")
            continue
            
    except Exception as e:
        print(f"Erro inesperado ao processar linha {i}: {e}")
        continue
# Dicionário para Quinzena
quinzena_dict = {}
fechamento['PK'] = fechamento['ROMANEIO'].astype(str) + "_" + fechamento['NF-E'].astype(str) + "_" + fechamento['CODPRODUTO'].astype(str)
for _, row in fechamento.iterrows():
    try:
        if pd.notna(row['QUINZENA']):
            quinzena_dict[row['PK']] = str(row['QUINZENA'])
    except:
        continue

# Dicionário para lookup de comissão por regra
comissao_regra_dict = {}
for _, row in fechamento.iterrows():
    try:
        key = (int(row['ROMANEIO']), int(row['NF-E']), int(row['CODPRODUTO']))
        comissao_regra_dict[key] = row['P.COM'] if pd.notna(row['P.COM']) else 0
    except:
        continue

# Dicionário para lookup por PK
fechamento_pk_dict = {}
coluna_desconto = 'Desconto verificado' if 'Desconto verificado' in fechamento.columns else 'DESCONTO'

for _, row in fechamento.iterrows():
    pk = str(row['ROMANEIO']) + "_" + str(row['NF-E']) + "_" + str(row['CODPRODUTO'])
    fechamento_pk_dict[pk] = {
        'ESCRITORIO': row['ESCRITORIO'] if 'ESCRITORIO' in fechamento.columns and pd.notna(row['ESCRITORIO']) else np.nan,
        'VLR ICMS': row['VLR ICMS'] if 'VLR ICMS' in fechamento.columns and pd.notna(row['VLR ICMS']) else np.nan,
        'PRECO VENDA': row['PRECO VENDA'] if 'PRECO VENDA' in fechamento.columns and pd.notna(row['PRECO VENDA']) else np.nan,
        'QUINZENA': row['QUINZENA'] if 'QUINZENA' in fechamento.columns and pd.notna(row['QUINZENA']) else ""
    }

# Dicionário para lookup por NF-E
fechamento_nf_dict = {}
for _, row in fechamento.iterrows():
    if pd.notna(row['NF-E']):
        fechamento_nf_dict[int(row['NF-E'])] = row['DESCRICAO'] if 'DESCRICAO' in fechamento.columns else ""

# Criar base_df
base_df = pd.DataFrame()

# Preencher colunas básicas - CORREÇÃO: QTDE deve ser o valor original
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
base_df['CODPRODUTO'] = fechamento_sem_cancelados['CODPRODUTO'].astype(str).str.strip()
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
# Preencher Quinzena
base_df['PK'] = base_df['OS'].astype(str) + "_" + base_df['NF-E'].astype(str) + "_" + base_df['CODPRODUTO'].astype(str)
base_df['Quinzena'] = base_df['PK'].map(lambda x: quinzena_dict.get(x, ""))
base_df['GRUPO'] = base_df['GRUPO'].fillna('VAREJO')



# 1. QTDE AJUSTADA
def calcular_qtde_ajustada(row):
    try:
        
        if row['QTDE REAL'] <= 0:
            return row['QTDE REAL']
        
        codproduto = str(row['CODPRODUTO']).strip() if pd.notna(row['CODPRODUTO']) else None
        data = row['DATA']
        
        if codproduto is None or data is None:
            return row['QTDE REAL']
            
        custo_info = custos_dict.get((codproduto, data), {})
        qtd = custo_info.get('QTD', 1)
        
        
        if qtd > 1:
            return row['QTDE'] * qtd
        else:
            return row['QTDE REAL']
    except:
        return row['QTDE REAL']

base_df['QTDE AJUSTADA'] = base_df.apply(calcular_qtde_ajustada, axis=1)

# 2. QTDE REAL2
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

base_df['QTDE REAL2'] = base_df.apply(calcular_qtde_real2, axis=1)

# Funções para buscar custo, frete e produção
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
    except Exception as e:
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
        frete = custo_info.get('FRETE', 0)
        
        return frete
    except:
        return 0

def buscar_producao(row):
    try:
        codproduto = str(row['CODPRODUTO']).strip() if pd.notna(row['CODPRODUTO']) else None
        data = row['DATA']
        
        if codproduto is None or data is None:
            return 0
            
        custo_info = custos_dict.get((codproduto, data), {})
        producao = custo_info.get('PRODUÇÃO', 0)
        
        return producao
    except:
        return 0

# Aplicar as funções
base_df['CUSTO'] = base_df.apply(buscar_custo, axis=1)
base_df['Frete'] = base_df.apply(buscar_frete, axis=1)
base_df['Produção'] = base_df.apply(buscar_producao, axis=1)

# 7. Escritório
if 'ESCRITORIO' in fechamento_sem_cancelados.columns:
    base_df['Escritório'] = fechamento_sem_cancelados['ESCRITORIO'].fillna(0) / 100
else:
    base_df['Escritório'] = 0

# MODIFICAÇÃO SOLICITADA: Substituir 4% por 3.5% na coluna Escritório
base_df['Escritório'] = base_df['Escritório'].apply(lambda x: 0.035 if abs(x - 0.04) < 0.001 else x)

# 8. P. Com
if 'P.COM' in fechamento_sem_cancelados.columns:
    base_df['P. Com'] = fechamento_sem_cancelados['P.COM'].fillna(0)
else:
    base_df['P. Com'] = 0

    
# 9. Desc. Valor - SOLUÇÃO ALTERNATIVA
# Mapear diretamente pelo índice para garantir correspondência
base_df['Desc Perc'] = 0  # Inicializar com zero

if 'DESCONTO' in fechamento_sem_cancelados.columns:
    for i, row in fechamento_sem_cancelados.iterrows():
        if i < len(base_df):  # Garantir que não ultrapasse o tamanho do base_df
            desconto_val = row['DESCONTO']
            if pd.notna(desconto_val) and str(desconto_val).strip() != '':
                try:
                    base_df.at[i, 'Desc Perc'] = float(str(desconto_val).replace(',', '.').strip()) / 100
                except:
                    base_df.at[i, 'Desc Perc'] = 0
            else:
                base_df.at[i, 'Desc Perc'] = 0

# Agora calcular o Desc. Valor
base_df['Desc. Valor'] = base_df.apply(
    lambda row: 0 if (row['CF'] == "DEV" or row['GRUPO'] == "TENDA") 
    else row['QTDE AJUSTADA'] * row['Preço Venda'] * row['Desc Perc'], axis=1
)

# 10. Fat. Bruto
base_df['Fat. Bruto'] = base_df.apply(
    lambda row: -row['QTDE AJUSTADA'] * row['Preço Venda'] if row['CF'] == "DEV"
    else row['QTDE AJUSTADA'] * row['Preço Venda'], axis=1
)

# 11. Aliq Icms
base_df['Aliq Icms'] = base_df.apply(
    lambda row: round(row['VL ICMS'] / row['Fat. Bruto'], 2) if (row['Fat. Bruto'] != 0 and pd.notna(row['VL ICMS'])) 
    else 0, axis=1
)

# Substituir infinitos por 0
base_df['Aliq Icms'] = base_df['Aliq Icms'].replace([np.inf, -np.inf], 0)

# Custo real
base_df['Custo real'] = base_df.apply(
    lambda row: 0 if (pd.isna(row['QTDE AJUSTADA']) or row['QTDE AJUSTADA'] <= 0 or 
                     pd.isna(row['CUSTO']) or pd.isna(row['Aliq Icms']))
    else row['CUSTO'] - (row['CUSTO'] * row['Aliq Icms']), axis=1
)

# 12. Fat Liquido
base_df['Fat Liquido'] = base_df.apply(
    lambda row: row['QTDE AJUSTADA'] * (row['Preço Venda'] - row['Preço Venda'] * row['Desc Perc']) if row['CF'] != "DEV"
    else -row['QTDE AJUSTADA'] * (row['Preço Venda'] - row['Preço Venda'] * row['Desc Perc']), axis=1
)

# 13. Aniversário
base_df['Aniversário'] = base_df.apply(
    lambda row: 0 if row['CF'] == "DEV" else row['Fat. Bruto'] * 0.01, axis=1
)

# 14. Comissão Kg
base_df['Comissão Kg'] = base_df.apply(
    lambda row: -(row['Preço Venda'] * row['P. Com']) if row['CF'] == "DEV" 
    else (row['Preço Venda'] * row['P. Com']), axis=1
)

# 15. Comissão Real
base_df['Comissão Real'] = base_df.apply(
    lambda row: row['Fat Liquido'] * row['P. Com'] if row['Preço Venda'] > 0 
    else -(row['Fat Liquido'] * row['P. Com']), axis=1
)

# 16. Coleta Esc
base_df['Coleta Esc'] = base_df['Fat. Bruto'] * base_df['Escritório']

# 17. Frete Real
base_df['Frete Real'] = base_df['QTDE REAL2'] * base_df['Frete']

# 18. comissão
base_df['comissão'] = base_df.apply(
    lambda row: row['Comissão Kg'] / row['Preço Venda'] if row['Preço Venda'] > 0
    else -row['Comissão Kg'] / row['Preço Venda'] if row['Preço Venda'] < 0
    else 0, axis=1
)

# 19. Escr.
base_df['Escr.'] = base_df.apply(
    lambda row: row['Coleta Esc'] / row['Fat. Bruto'] if row['Fat. Bruto'] != 0
    else 0, axis=1
)

# 20. frete
base_df['frete'] = base_df.apply(
    lambda row: row['Frete Real'] / row['Fat. Bruto'] if row['Fat. Bruto'] != 0
    else 0, axis=1
)

# 21. TP
base_df['TP'] = base_df.apply(
    lambda row: "BIG BACON" if str(row['CODPRODUTO']) == "700"
    else "REFFINATO" if row['GRUPO PRODUTO'] in ["SALGADOS SUINOS A GRANEL", "SALGADOS SUINOS EMBALADOS"]
    else "MIX", axis=1
)

# 22. CANC
base_df['CANC'] = base_df['NF-E'].apply(lambda x: "SIM" if x in notas_canceladas else "")

# 23. Armazenagem
base_df['Armazenagem'] = base_df.apply(
    lambda row: (row['Fat. Bruto'] * row['P. Com']) / row['QTDE AJUSTADA'] if row['QTDE AJUSTADA'] != 0
    else 0, axis=1
)

# 24. Comissão por Regra
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

# 25. Coluna2
base_df['Coluna2'] = base_df['Comissão por Regra'] == base_df['Comissão Kg']

# 26. FRETE - LUC/PREJ
base_df['FRETE - LUC/PREJ'] = base_df['QTDE AJUSTADA'] * base_df['Frete']

# 27. DESC FEC
def buscar_desc_fec(row):
    try:
        pk = str(row['OS']) + "_" + str(row['NF-E']) + "_" + str(row['CODPRODUTO'])
        return fechamento_pk_dict.get(pk, {}).get('DESCONTO', np.nan)
    except:
        return np.nan

base_df['DESC FEC'] = base_df.apply(buscar_desc_fec, axis=1)

# 28. ESC FEC
def buscar_esc_fec(row):
    try:
        pk = str(row['OS']) + "_" + str(row['NF-E']) + "_" + str(row['CODPRODUTO'])
        return fechamento_pk_dict.get(pk, {}).get('ESCRITORIO', np.nan)
    except:
        return np.nan

base_df['ESC FEC'] = base_df.apply(buscar_esc_fec, axis=1)

# 29. ICMS FEC
def buscar_icms_fec(row):
    try:
        pk = str(row['OS']) + "_" + str(row['NF-E']) + "_" + str(row['CODPRODUTO'])
        return fechamento_pk_dict.get(pk, {}).get('VLR ICMS', np.nan)
    except:
        return np.nan

base_df['ICMS FEC'] = base_df.apply(buscar_icms_fec, axis=1)

# 30. PRC VEND FEV
def buscar_prc_vend_fev(row):
    try:
        pk = str(row['OS']) + "_" + str(row['NF-E']) + "_" + str(row['CODPRODUTO'])
        return fechamento_pk_dict.get(pk, {}).get('PRECO VENDA', np.nan)
    except:
        return np.nan

base_df['PRC VEND FEV'] = base_df.apply(buscar_prc_vend_fev, axis=1)

# 31. DESC
base_df['DESC'] = base_df.apply(
    lambda row: (row['DESC FEC'] / 100) == row['Desc Perc'] if pd.notna(row['DESC FEC']) else False, axis=1
)

# 32. ESC
base_df['ESC'] = base_df.apply(
    lambda row: (row['ESC FEC'] / 100) == row['Escritório'] if pd.notna(row['ESC FEC']) else False, axis=1
)

# 33. ICMS
base_df['ICMS'] = base_df.apply(
    lambda row: row['ICMS FEC'] == row['VL ICMS'] if pd.notna(row['ICMS FEC']) else False, axis=1
)

# 34. PRC VEND
base_df['PRC VEND'] = base_df.apply(
    lambda row: row['PRC VEND FEV'] == row['Preço Venda'] if pd.notna(row['PRC VEND FEV']) else False, axis=1
)

# 35. DESCRIÇÃO_1
base_df['DESCRIÇÃO_1'] = base_df['NF-E'].apply(lambda x: fechamento_nf_dict.get(int(x), '') if pd.notna(x) else '')

# 36. MOV ENC
base_df['MOV ENC'] = base_df.apply(
    lambda row: "ENCONTRADO" if any([str(row['OS']) == str(venda[0]) and str(row['NF-E']) == str(venda[1]) for venda in vendas_var])
    else "NÃO ENCONTRADO", axis=1
)

# 37. CUST + IMP
base_df['CUST + IMP'] = base_df['Custo real'] * base_df['QTDE AJUSTADA']

# 38. CUST PROD
base_df['CUST PROD'] = base_df['Custo real'] * base_df['QTDE AJUSTADA']

# 39. COM BRUTA
base_df['COM BRUTA'] = base_df['P. Com'] * base_df['Fat. Bruto']

# 40. Coluna1
base_df['Coluna1'] = (round(base_df['COM BRUTA'], 2) == round(base_df['Comissão Real'], 2))

# 41. Custo divergente
base_df['Custo divergente'] = base_df.apply(
    lambda row: "Só constando" if (row['QTDE'] > 0 and row['CUSTO EM SISTEMA'] == row['CUSTO']) else "", axis=1
)

# 42. Lucro / Prej.
base_df['Lucro / Prej.'] = base_df['Fat Liquido'] - base_df['CUST + IMP']

# 43. Margem
base_df['Margem'] = base_df.apply(
    lambda row: (row['Lucro / Prej.'] / row['Fat Liquido'] * 100) if row['Fat Liquido'] != 0 else 0, axis=1
)

# 44. INCL.
base_df['INCL.'] = ""

# 45. DESCRIÇÃO_2
base_df['DESCRIÇÃO_2'] = ""

# Reordenar colunas na ordem solicitada
colunas_ordenadas = [
    'CF', 'RAZAO', 'FANTASIA', 'GRUPO', 'OS', 'NF-E', 'CF_NF', 'DATA', 'VENDEDOR',
    'CODPRODUTO', 'GRUPO PRODUTO', 'DESCRICAO', 'QTDE', 'QTDE REAL', 'CUSTO EM SISTEMA',
    'QTDE AJUSTADA', 'QTDE REAL2', 'CUSTO', 'Custo real', 'Frete', 'Produção',
    'Escritório', 'Comissão Kg', 'P. Com', 'Aniversário', 'Val Pis', 'VLRCOFINS',
    'IRPJ', 'CSLL', 'VL ICMS', 'Aliq Icms', 'Desc Perc', 'Desc. Valor', 'Preço Venda',
    'Fat Liquido', 'Fat. Bruto', 'Lucro / Prej.', 'Margem', 'Quinzena', 'Comissão Real',
    'Coleta Esc', 'Frete Real', 'INCL.', 'comissão', 'Escr.', 'frete', 'Custo divergente',
    'TP', 'CANC', 'Armazenagem', 'Comissão por Regra', 'PK', 'Coluna2', 'FRETE - LUC/PREJ',
    'CUST + IMP', 'CUST PROD', 'COM BRUTA', 'DESC FEC', 'ESC FEC', 'ICMS FEC', 'PRC VEND FEV',
    'DESC', 'ESC', 'ICMS', 'PRC VEND', 'Coluna1', 'DESCRIÇÃO_1', 'DESCRIÇÃO_2'
]

# Manter apenas colunas que existem no DataFrame
colunas_existentes = [col for col in colunas_ordenadas if col in base_df.columns]
base_df = base_df[colunas_existentes]

# Substituir NaN por string vazia para melhor visualização
base_df = base_df.fillna("")

# Criar arquivo Excel
output_path = f"C:\\Users\\win11\\Downloads\\Margem_{data_nome_arquivo}.xlsx"

with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    base_df.to_excel(writer, sheet_name='base (3,5%)', index=False)
    ofertas_vog.to_excel(writer, sheet_name='OFERTAS_VOG', index=False)
    custos_produtos.to_excel(writer, sheet_name='PRECOS', index=False)
    cancelados.to_excel(writer, sheet_name='Base_cancelamento', index=False)
    devolucoes.to_excel(writer, sheet_name='Base_movimentacao', index=False)
    fechamento.to_excel(writer, sheet_name='Base_Fechamento', index=False)

# Salvar também como JSON
json_path = f"C:\\Users\\win11\\Downloads\\Margem_{data_nome_arquivo}.json"

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
    raise TypeError(f"Type {type(obj)} not serializable")

with open(json_path, 'w', encoding='utf-8') as f:
    json.dump(base_df.to_dict(orient='records'), f, ensure_ascii=False, indent=4, default=default_serializer)

print(f"\nArquivo Excel salvo em: {output_path}")
print(f"Arquivo JSON salvo em: {json_path}")
print("Processo concluído!\n")