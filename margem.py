import pandas as pd
import numpy as np
from datetime import datetime, date
import warnings
import json
import os
from collections import defaultdict

warnings.filterwarnings('ignore')

# Data fixa para o nome do arquivo
data_nome_arquivo = "050825"

# Carregar os arquivos
print("Carregando arquivos...")

# Fechamento.csv
fechamento = pd.read_csv(r"C:\Users\win11\Downloads\fechamento.csv", sep=';', encoding='utf-8', decimal=',')
print(f"Fechamento carregado: {len(fechamento)} linhas")

# Cancelados.csv (pula as 2 primeiras linhas)
cancelados = pd.read_csv(r"C:\Users\win11\Downloads\cancelados.csv", sep=';', encoding='utf-8', decimal=',', skiprows=2)
print(f"Cancelados carregado: {len(cancelados)} linhas")

# Devoluções.csv
devolucoes = pd.read_csv(r"C:\Users\win11\Downloads\movimentação.csv", sep=';', encoding='utf-8', decimal=',')
print(f"Devoluções carregado: {len(devolucoes)} linhas")

# Custos de produtos - Agosto.xlsx
custos_produtos = pd.read_excel(r"C:\Users\win11\Downloads\Custos de produtos - Julho.xlsx", sheet_name='Base')
print(f"Custos produtos carregado: {len(custos_produtos)} linhas")

# OFERTAS_VOG.xlsx
ofertas_vog = pd.read_excel(r"C:\Users\win11\Downloads\OFERTAS_VOG.xlsx")
print(f"Ofertas VOG carregado: {len(ofertas_vog)} linhas")

# Converter colunas numéricas para o tipo correto
print("Convertendo tipos de dados...")

# Converter colunas numéricas em custos_produtos
numeric_columns_custos = ['PRODUTO', 'PCS', 'KGS', 'CUSTO', 'TOTAL', 'FRETE', 'PRODUÇÃO']
for col in numeric_columns_custos:
    if col in custos_produtos.columns:
        custos_produtos[col] = custos_produtos[col].astype(str)
        custos_produtos[col] = custos_produtos[col].str.replace('.', '', regex=False)
        custos_produtos[col] = custos_produtos[col].str.replace(',', '.', regex=False)
        custos_produtos[col] = pd.to_numeric(custos_produtos[col], errors='coerce')

# Renomear colunas de custos_produtos para facilitar o lookup
custos_produtos.rename(columns={
    'PRODUTO': 'CODPRODUTO',
    'PCS': 'QTD',
    'KGS': 'PESO',
    'TOTAL': 'CUSTO_TOTAL'
}, inplace=True)

# Converter DATA para datetime em custos_produtos
custos_produtos['DATA'] = pd.to_datetime(custos_produtos['DATA'], errors='coerce')

# Converter colunas numéricas em fechamento
numeric_columns_fechamento = ['ROMANEIO', 'NF-E', 'CF_NF', 'CODPRODUTO', 'QTDE', 'QTDE REAL', 'CUSTO', 
                             'FRETE', 'PRODUCAO', 'ESCRITORIO', 'P.COM', 'ANIVERSARIO', 'VLR PIS', 
                             'VLR COFINS', 'IRPJ', 'CSLL', 'VLR ICMS', 'ALIQ ICMS', 'DESCONTO', 
                             'VLR DESCONTO', 'PRECO VENDA', 'FAT LIQUIDO', 'FAT BRUTO', 'LUCRO', 'MARGEM', 
                             'QUINZENA', 'QTD POR EMB', 'FATOR DE CONVERSAO']

for col in numeric_columns_fechamento:
    if col in fechamento.columns:
        fechamento[col] = fechamento[col].astype(str)
        fechamento[col] = fechamento[col].str.replace('.', '', regex=False)
        fechamento[col] = fechamento[col].str.replace(',', '.', regex=False)
        fechamento[col] = pd.to_numeric(fechamento[col], errors='coerce')

# Criar variáveis conforme especificado
print("Criando variáveis...")

# 1. Notas canceladas
notas_canceladas = cancelados['NUMERO'].tolist()
print(f"Notas canceladas: {len(notas_canceladas)} registros")

# 2. Devoluções (DESCRICAO = "DEV VENDA C/ FIN S/ EST" ou HISTORICO = "68")
devolucoes_filtro = devolucoes[
    (devolucoes['DESCRICAO'] == "DEV VENDA C/ FIN S/ EST") | 
    (devolucoes['HISTORICO'] == "68")
]
devolucoes_var = devolucoes_filtro[['ROMANEIO', 'NOTA FISCAL']].values.tolist()
print(f"Devoluções: {len(devolucoes_var)} registros")

# 3. Vendas (DESCRICAO = "VENDA" ou HISTORICO = "51")
vendas_filtro = devolucoes[
    (devolucoes['DESCRICAO'] == "VENDA") | 
    (devolucoes['HISTORICO'] == "51")
]
vendas_var = vendas_filtro[['ROMANEIO', 'NOTA FISCAL']].values.tolist()
print(f"Vendas: {len(vendas_var)} registros")

# Filtrar fechamento removendo notas canceladas
fechamento_sem_cancelados = fechamento[~fechamento['NF-E'].isin(notas_canceladas)].copy()
print(f"Fechamento sem cancelados: {len(fechamento_sem_cancelados)} linhas")

# Criar dicionários para lookup rápido
print("Criando dicionários para lookup...")

# Dicionário para custos_produtos (por CODPRODUTO)
custos_dict = {}
for _, row in custos_produtos.iterrows():
    if pd.notna(row['CODPRODUTO']):
        try:
            codproduto = int(row['CODPRODUTO'])
            custos_dict[codproduto] = {
                'QTD': float(row['QTD']) if pd.notna(row['QTD']) else np.nan,
                'PESO': float(row['PESO']) if pd.notna(row['PESO']) else np.nan,
                'CUSTO': float(row['CUSTO']) if pd.notna(row['CUSTO']) else np.nan,
                'FRETE': float(row['FRETE']) if pd.notna(row['FRETE']) else 0,
                'PRODUÇÃO': float(row['PRODUÇÃO']) if pd.notna(row['PRODUÇÃO']) else 0
            }
        except:
            continue

print(f"Dicionário de custos criado com {len(custos_dict)} entradas")

# Dicionário para lookup de comissão por regra
comissao_regra_dict = {}
for _, row in fechamento.iterrows():
    try:
        key = (int(row['ROMANEIO']), int(row['NF-E']), int(row['CODPRODUTO']))
        comissao_regra_dict[key] = row['P.COM']
    except:
        continue

# Dicionário para lookup por PK
fechamento_pk_dict = {}
fechamento['PK'] = fechamento['ROMANEIO'].astype(str) + "_" + fechamento['NF-E'].astype(str) + "_" + fechamento['CODPRODUTO'].astype(str)
for _, row in fechamento.iterrows():
    fechamento_pk_dict[row['PK']] = {
        'DESCONTO': row['DESCONTO'] if 'DESCONTO' in fechamento.columns and pd.notna(row['DESCONTO']) else np.nan,
        'ESCRITORIO': row['ESCRITORIO'] if 'ESCRITORIO' in fechamento.columns and pd.notna(row['ESCRITORIO']) else np.nan,
        'VLR ICMS': row['VLR ICMS'] if 'VLR ICMS' in fechamento.columns and pd.notna(row['VLR ICMS']) else np.nan,
        'PRECO VENDA': row['PRECO VENDA'] if 'PRECO VENDA' in fechamento.columns and pd.notna(row['PRECO VENDA']) else np.nan
    }

# Dicionário para lookup por NF-E
fechamento_nf_dict = {}
for _, row in fechamento.iterrows():
    if pd.notna(row['NF-E']):
        fechamento_nf_dict[int(row['NF-E'])] = row['DESCRICAO'] if 'DESCRICAO' in fechamento.columns else ""

# Preparar dados para a tabela base (3,5%)
print("Preparando tabela base (3,5%)...")

# Criar DataFrame base
base_df = pd.DataFrame()

# Preencher colunas básicas primeiro
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
base_df['CODPRODUTO'] = fechamento_sem_cancelados['CODPRODUTO']
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
base_df['Aliq Icms'] = fechamento_sem_cancelados['ALIQ ICMS'] / 100 if 'ALIQ ICMS' in fechamento_sem_cancelados.columns else 0
base_df['Desc Perc'] = fechamento_sem_cancelados['DESCONTO'].fillna(0) / 100 if 'DESCONTO' in fechamento_sem_cancelados.columns else 0
base_df['Preço Venda'] = fechamento_sem_cancelados['PRECO VENDA'] if 'PRECO VENDA' in fechamento_sem_cancelados.columns else 0
base_df['Quinzena'] = fechamento_sem_cancelados['QUINZENA'] if 'QUINZENA' in fechamento_sem_cancelados.columns else ""

# Aplicar cálculos em ordem correta
print("Calculando colunas...")

# 1. QTDE AJUSTADA
def calcular_qtde_ajustada(row):
    try:
        if row['QTDE REAL'] <= 0:
            return row['QTDE REAL']
        
        codproduto = int(row['CODPRODUTO']) if pd.notna(row['CODPRODUTO']) else None
        if codproduto is None:
            return row['QTDE REAL']
            
        custo_info = custos_dict.get(codproduto, {})
        peso = custo_info.get('PESO', np.nan)
        
        if pd.notna(peso) and peso > 1:
            return row['QTDE'] * peso
        else:
            return row['QTDE REAL']
    except:
        return row['QTDE REAL']

base_df['QTDE AJUSTADA'] = base_df.apply(calcular_qtde_ajustada, axis=1)

# 2. QTDE REAL2
def calcular_qtde_real2(row):
    try:
        codproduto = int(row['CODPRODUTO']) if pd.notna(row['CODPRODUTO']) else None
        if codproduto is None:
            return np.nan
            
        custo_info = custos_dict.get(codproduto, {})
        peso = custo_info.get('PESO', np.nan)
        
        if pd.isna(peso):
            return np.nan
            
        if row['QTDE REAL'] < 0:
            return -row['QTDE AJUSTADA'] * peso
        else:
            return row['QTDE AJUSTADA'] * peso
    except:
        return np.nan

base_df['QTDE REAL2'] = base_df.apply(calcular_qtde_real2, axis=1)

# 3. CUSTO
def buscar_custo(row):
    try:
        codproduto = int(row['CODPRODUTO']) if pd.notna(row['CODPRODUTO']) else None
        if codproduto is None:
            return np.nan
            
        custo_info = custos_dict.get(codproduto, {})
        custo = custo_info.get('CUSTO', np.nan)
        
        return custo
    except:
        return np.nan

base_df['CUSTO'] = base_df.apply(buscar_custo, axis=1)

# 4. Custo real
base_df['Custo real'] = base_df.apply(
    lambda row: 0 if (pd.isna(row['QTDE AJUSTADA']) or row['QTDE AJUSTADA'] <= 0 or 
                     pd.isna(row['CUSTO']) or pd.isna(row['Aliq Icms']))
    else row['CUSTO'] - (row['CUSTO'] * row['Aliq Icms']), axis=1
)

# 5. Frete
def buscar_frete(row):
    if row['FANTASIA'] in ["PASSOS ALIMENTOS LTDA", "AGELLE ARMAZEM E LOGISTICA LTDA", 
                           "GEMEOS REORESENTACOES", "REAL DISTRIBUIDORA"]:
        return 0
        
    try:
        codproduto = int(row['CODPRODUTO']) if pd.notna(row['CODPRODUTO']) else None
        if codproduto is None:
            return 0
            
        custo_info = custos_dict.get(codproduto, {})
        frete = custo_info.get('FRETE', 0)
        
        return frete
    except:
        return 0

base_df['Frete'] = base_df.apply(buscar_frete, axis=1)

# 6. Produção
def buscar_producao(row):
    try:
        codproduto = int(row['CODPRODUTO']) if pd.notna(row['CODPRODUTO']) else None
        if codproduto is None:
            return 0
            
        custo_info = custos_dict.get(codproduto, {})
        producao = custo_info.get('PRODUÇÃO', 0)
        
        return producao
    except:
        return 0

base_df['Produção'] = base_df.apply(buscar_producao, axis=1)

# 7. Escritório
if 'ESCRITORIO' in fechamento_sem_cancelados.columns:
    base_df['Escritório'] = fechamento_sem_cancelados['ESCRITORIO'].fillna(0) / 100
else:
    base_df['Escritório'] = 0

# 8. P. Com
if 'P.COM' in fechamento_sem_cancelados.columns:
    base_df['P. Com'] = fechamento_sem_cancelados['P.COM'].fillna(0)
else:
    base_df['P. Com'] = 0

# 9. Desc. Valor
base_df['Desc. Valor'] = base_df.apply(
    lambda row: 0 if (row['CF'] == "DEV" or row['GRUPO'] == "TENDA") 
    else row['QTDE AJUSTADA'] * row['Preço Venda'] * row['Desc Perc'], axis=1
)

# 10. Fat. Bruto
base_df['Fat. Bruto'] = base_df.apply(
    lambda row: -row['QTDE AJUSTADA'] * row['Preço Venda'] if row['CF'] == "DEV"
    else row['QTDE AJUSTADA'] * row['Preço Venda'], axis=1
)

# 11. Fat Liquido
base_df['Fat Liquido'] = base_df.apply(
    lambda row: row['QTDE AJUSTADA'] * (row['Preço Venda'] - row['Preço Venda'] * row['Desc Perc']) if row['CF'] != "DEV"
    else -row['QTDE AJUSTADA'] * (row['Preço Venda'] - row['Preço Venda'] * row['Desc Perc']), axis=1
)

# 12. Aniversário
base_df['Aniversário'] = base_df.apply(
    lambda row: 0 if row['CF'] == "DEV" else row['Fat. Bruto'] * 0.01, axis=1
)

# 13. Comissão Kg
base_df['Comissão Kg'] = base_df.apply(
    lambda row: -(row['Preço Venda'] * row['P. Com']) if row['CF'] == "DEV" 
    else (row['Preço Venda'] * row['P. Com']), axis=1
)

# 14. Comissão Real
base_df['Comissão Real'] = base_df.apply(
    lambda row: row['Fat Liquido'] * row['P. Com'] if row['Preço Venda'] > 0 
    else -(row['Fat Liquido'] * row['P. Com']), axis=1
)

# 15. Coleta Esc
base_df['Coleta Esc'] = base_df['Fat. Bruto'] * base_df['Escritório']

# 16. Frete Real
base_df['Frete Real'] = base_df['QTDE REAL2'] * base_df['Frete']

# 17. comissão
base_df['comissão'] = np.where(
    base_df['Preço Venda'] > 0,
    base_df['Comissão Kg'] / base_df['Preço Venda'],
    -base_df['Comissão Kg'] / base_df['Preço Venda']
)

# 18. Escr.
base_df['Escr.'] = base_df['Coleta Esc'] / base_df['Fat. Bruto']

# 19. frete
base_df['frete'] = base_df['Frete Real'] / base_df['Fat. Bruto']

# 20. TP
base_df['TP'] = base_df.apply(
    lambda row: "BIG BACON" if row['CODPRODUTO'] == 700
    else "REFFINATO" if row['GRUPO PRODUTO'] in ["SALGADOS SUINOS A GRANEL", "SALGADOS SUINOS EMBALADOS"]
    else "MIX", axis=1
)

# 21. CANC
base_df['CANC'] = base_df['NF-E'].apply(lambda x: "SIM" if x in notas_canceladas else "")

# 22. Armazenagem
base_df['Armazenagem'] = (base_df['Fat. Bruto'] * base_df['P. Com']) / base_df['QTDE AJUSTADA']

# 23. Comissão por Regra
base_df['Comissão por Regra'] = base_df.apply(
    lambda row: comissao_regra_dict.get((int(row['OS']), int(row['NF-E']), int(row['CODPRODUTO'])), 0)
    if pd.notna(row['OS']) and pd.notna(row['NF-E']) and pd.notna(row['CODPRODUTO'])
    else 0, axis=1
)

# 24. PK
base_df['PK'] = base_df['OS'].astype(str) + "_" + base_df['NF-E'].astype(str) + "_" + base_df['CODPRODUTO'].astype(str)

# 25. Coluna2
base_df['Coluna2'] = base_df['Comissão por Regra'] == base_df['Comissão Kg']

# 26. FRETE - LUC/PREJ
base_df['FRETE - LUC/PREJ'] = base_df['QTDE AJUSTADA'] * base_df['Frete']

# 27. DESC FEC
base_df['DESC FEC'] = base_df['PK'].apply(lambda x: fechamento_pk_dict.get(x, {}).get('DESCONTO', 'x'))

# 28. ESC FEC
base_df['ESC FEC'] = base_df['PK'].apply(lambda x: fechamento_pk_dict.get(x, {}).get('ESCRITORIO', 'x'))

# 29. ICMS FEC
base_df['ICMS FEC'] = base_df['PK'].apply(lambda x: fechamento_pk_dict.get(x, {}).get('VLR ICMS', 'x'))

# 30. PRC VEND FEV
base_df['PRC VEND FEV'] = base_df['PK'].apply(lambda x: fechamento_pk_dict.get(x, {}).get('PRECO VENDA', 'x'))

# 31. DESC
base_df['DESC'] = base_df.apply(
    lambda row: (row['DESC FEC'] / 100) == row['Desc Perc'] if row['DESC FEC'] != 'x' else False, axis=1
)

# 32. ESC
base_df['ESC'] = base_df.apply(
    lambda row: (row['ESC FEC'] / 100) == row['Escritório'] if row['ESC FEC'] != 'x' else False, axis=1
)

# 33. ICMS
base_df['ICMS'] = base_df['ICMS FEC'] == base_df['VL ICMS']

# 34. PRC VEND
base_df['PRC VEND'] = base_df['PRC VEND FEV'] == base_df['Preço Venda']

# 35. DESCRIÇÃO_1
base_df['DESCRIÇÃO_1'] = base_df['NF-E'].apply(lambda x: fechamento_nf_dict.get(x, ''))

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
base_df['Custo divergente'] = np.where(
    base_df['QTDE'] > 0,
    np.where(base_df['CUSTO EM SISTEMA'] == base_df['CUSTO'], "Só constando", ""),
    ""
)

# 42. Lucro / Prej.
base_df['Lucro / Prej.'] = base_df['Fat Liquido'] - base_df['CUST + IMP']

# 43. Margem
base_df['Margem'] = np.where(
    base_df['Fat Liquido'] != 0,
    base_df['Lucro / Prej.'] / base_df['Fat Liquido'] * 100,
    0
)

# 44. INCL.
base_df['INCL.'] = ""

# 45. DESCRIÇÃO_2
base_df['DESCRIÇÃO_2'] = ""

# Reordenar colunas para manter a ordem desejada
colunas_ordenadas = [
    'CF', 'RAZAO', 'FANTASIA', 'GRUPO', 'OS', 'NF-E', 'CF_NF', 'DATA', 'VENDEDOR', 
    'CODPRODUTO', 'GRUPO PRODUTO', 'DESCRICAO', 'QTDE', 'QTDE REAL', 'CUSTO EM SISTEMA', 
    'Val Pis', 'VLRCOFINS', 'IRPJ', 'CSLL', 'VL ICMS', 'Aliq Icms', 'Desc Perc', 
    'Desc. Valor', 'Preço Venda', 'Fat Liquido', 'Fat. Bruto', 'Lucro / Prej.', 'Margem', 
    'Quinzena', 'QTDE AJUSTADA', 'QTDE REAL2', 'CUSTO', 'Custo real', 'Frete', 'Produção', 
    'Escritório', 'P. Com', 'Aniversário', 'Comissão Kg', 'Comissão Real', 'Coleta Esc', 
    'Frete Real', 'comissão', 'Escr.', 'frete', 'TP', 'CANC', 'Armazenagem', 
    'Comissão por Regra', 'PK', 'Coluna2', 'FRETE - LUC/PREJ', 'DESC FEC', 'ESC FEC', 
    'ICMS FEC', 'PRC VEND FEV', 'DESC', 'ESC', 'ICMS', 'PRC VEND', 'DESCRIÇÃO_1', 
    'MOV ENC', 'INCL.', 'Custo divergente', 'CUST + IMP', 'CUST PROD', 'COM BRUTA', 
    'Coluna1', 'DESCRIÇÃO_2'
]

base_df = base_df[colunas_ordenadas]

# Criar arquivo Excel
print("Criando arquivo Excel...")
output_path = f"C:\\Users\\win11\\Downloads\\margem_{data_nome_arquivo}.xlsx"

with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    base_df.to_excel(writer, sheet_name='base (3,5%)', index=False)
    ofertas_vog.to_excel(writer, sheet_name='OFERTAS_VOG', index=False)
    custos_produtos.to_excel(writer, sheet_name='PRECOS', index=False)
    cancelados.to_excel(writer, sheet_name='Base_cancelamento', index=False)
    devolucoes.to_excel(writer, sheet_name='Base_movimentacao', index=False)
    fechamento.to_excel(writer, sheet_name='Base_Fechamento', index=False)

# Salvar também como JSON
json_path = f"C:\\Users\\win11\\Downloads\\margem_{data_nome_arquivo}.json"

# Converter para JSON com tratamento de valores NaN
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
        return ""
    raise TypeError(f"Type {type(obj)} not serializable")

with open(json_path, 'w', encoding='utf-8') as f:
    json.dump(base_df.to_dict(orient='records'), f, ensure_ascii=False, indent=4, default=default_serializer)

print(f"Arquivo Excel salvo em: {output_path}")
print(f"Arquivo JSON salvo em: {json_path}")
print("Processo concluído!")