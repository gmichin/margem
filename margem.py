import pandas as pd
import numpy as np
from datetime import datetime, date
import warnings
import json
import os
from collections import defaultdict
import re
import os

# Limpar console
os.system('cls' if os.name == 'nt' else 'clear')

warnings.filterwarnings('ignore')

# Data fixa para o nome do arquivo
data_nome_arquivo = "050825"

# Fechamento.csv - CORREÇÃO: especificar dtype para evitar problemas com decimais
fechamento = pd.read_csv(r"C:\Users\win11\Downloads\fechamento.csv", sep=';', encoding='utf-8', decimal=',', thousands='.')
print(f"Fechamento carregado: {len(fechamento)} linhas")

# Cancelados.csv (pula as 2 primeiras linhas)
cancelados = pd.read_csv(r"C:\Users\win11\Downloads\cancelados.csv", sep=';', encoding='utf-8', decimal=',', thousands='.', skiprows=2)
print(f"Cancelados carregado: {len(cancelados)} linhas")

# Devoluções.csv
devolucoes = pd.read_csv(r"C:\Users\win11\Downloads\movimentação.csv", sep=';', encoding='utf-8', decimal=',', thousands='.')
print(f"Devoluções carregado: {len(devolucoes)} linhas")

# Função para converter valores brasileiros para float
def converter_valor_brasileiro(valor):
    if pd.isna(valor) or valor == '':
        return 0.0
    try:
        # Se já for numérico, retorna direto
        if isinstance(valor, (int, float)):
            return float(valor)
        
        # Remove pontos de milhar e substitui vírgula decimal por ponto
        valor_str = str(valor).strip()
        valor_str = valor_str.replace('.', '').replace(',', '.')
        
        # Remove caracteres não numéricos exceto ponto e sinal negativo
        valor_str = re.sub(r'[^\d\.\-]', '', valor_str)
        
        return float(valor_str)
    except:
        return 0.0

# Custos de produtos - Agosto.xlsx
custos_produtos = pd.read_excel(r"C:\Users\win11\Downloads\Custos de produtos - Julho.xlsx", sheet_name='Base')
print(f"Custos produtos carregado: {len(custos_produtos)} linhas")
colunas_numericas = ['PCS', 'KGS', 'CUSTO', 'TOTAL', 'PRODUÇÃO', 'FRETE']

for coluna in colunas_numericas:
    if coluna in custos_produtos.columns:
        custos_produtos[coluna] = custos_produtos[coluna].apply(converter_valor_brasileiro)

# OFERTAS_VOG.xlsx
ofertas_vog = pd.read_excel(r"C:\Users\win11\Downloads\OFERTAS_VOG.xlsx")

# Renomear colunas de custos_produtos para facilitar o lookup
custos_produtos.rename(columns={
    'PRODUTO': 'CODPRODUTO',
    'PCS': 'QTD',
    'KGS': 'PESO',
    'TOTAL': 'CUSTO_TOTAL'
}, inplace=True)

# CORREÇÃO: Converter DATA para datetime em custos_produtos usando o formato dia/mês/ano
custos_produtos['DATA'] = pd.to_datetime(custos_produtos['DATA'], format='%d/%m/%Y', errors='coerce')

# CORREÇÃO: Converter colunas numéricas em fechamento mantendo valores decimais corretos
numeric_columns_fechamento = ['ROMANEIO', 'NF-E', 'CF_NF', 'CODPRODUTO', 'QTDE', 'QTDE REAL', 'CUSTO', 
                             'FRETE', 'PRODUCAO', 'ESCRITORIO', 'P.COM', 'ANIVERSARIO', 'VLR PIS', 
                             'VLR COFINS', 'IRPJ', 'CSLL', 'VLR ICMS', 'ALIQ ICMS', 'DESCONTO', 
                             'VLR DESCONTO', 'PRECO VENDA', 'FAT LIQUIDO', 'FAT BRUTO', 'LUCRO', 'MARGEM', 
                             'QTD POR EMB', 'FATOR DE CONVERSAO']

for col in numeric_columns_fechamento:
    if col in fechamento.columns:
        if col == 'DESCONTO':  # Tratamento especial para a coluna DESCONTO
            # Converter usando a mesma lógica da função converter_valor_brasileiro
            fechamento[col] = fechamento[col].apply(converter_valor_brasileiro)
            
            # Dividir por 100 para obter a porcentagem correta
            fechamento[col] = fechamento[col] / 100
        else:
            # Para outras colunas, usar conversão padrão
            fechamento[col] = pd.to_numeric(fechamento[col], errors='coerce')

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

# CORREÇÃO: Dicionário para custos_produtos (por CODPRODUTO e DATA)
custos_dict = {}
custos_produtos_sem_data = 0
custos_produtos_sem_codigo = 0

for _, row in custos_produtos.iterrows():
    if pd.isna(row['DATA']):
        custos_produtos_sem_data += 1
        continue
        
    if pd.isna(row['CODPRODUTO']):
        custos_produtos_sem_codigo += 1
        continue
        
    try:
        codproduto = int(row['CODPRODUTO'])
        data_key = row['DATA'].date()  # Garantir que é apenas a data (sem hora)
        
        # Usar valores convertidos corretamente
        custo_val = float(row['CUSTO']) if pd.notna(row['CUSTO']) else 0
        peso_val = float(row['PESO']) if pd.notna(row['PESO']) else 1
        producao_val = float(row['PRODUÇÃO']) if pd.notna(row['PRODUÇÃO']) else 0
        frete_val = float(row['FRETE']) if pd.notna(row['FRETE']) else 0
        
        custos_dict[(codproduto, data_key)] = {
            'QTD': float(row['QTD']) if pd.notna(row['QTD']) else 0,
            'PESO': peso_val,
            'CUSTO': custo_val,
            'FRETE': frete_val,
            'PRODUÇÃO': producao_val
        }
    except Exception as e:
        continue

# Dicionário para Quinzena - CORREÇÃO: usar PK completa (ROMANEIO_NF-E_CODPRODUTO)
quinzena_dict = {}
fechamento['PK'] = fechamento['ROMANEIO'].astype(str) + "_" + fechamento['NF-E'].astype(str) + "_" + fechamento['CODPRODUTO'].astype(str)
for _, row in fechamento.iterrows():
    try:
        if pd.notna(row['QUINZENA']):
            quinzena_dict[row['PK']] = str(row['QUINZENA'])  # Garantir que seja string
    except:
        continue

# Dicionário para lookup de comissão por regra (usando múltiplas condições)
comissao_regra_dict = {}
for _, row in fechamento.iterrows():
    try:
        key = (int(row['ROMANEIO']), int(row['NF-E']), int(row['CODPRODUTO']))
        comissao_regra_dict[key] = row['P.COM'] if pd.notna(row['P.COM']) else 0
    except:
        continue

# Dicionário para lookup por PK
fechamento_pk_dict = {}

# Verificar se a coluna 'Desconto verificado' existe, caso contrário usar 'DESCONTO'
coluna_desconto = 'Desconto verificado' if 'Desconto verificado' in fechamento.columns else 'DESCONTO'

for _, row in fechamento.iterrows():
    pk = str(row['ROMANEIO']) + "_" + str(row['NF-E']) + "_" + str(row['CODPRODUTO'])
    fechamento_pk_dict[pk] = {
        'DESCONTO': row[coluna_desconto] if pd.notna(row[coluna_desconto]) else np.nan,
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
base_df['Desc Perc'] = fechamento_sem_cancelados['DESCONTO'].fillna(0) if 'DESCONTO' in fechamento_sem_cancelados.columns else 0
base_df['Preço Venda'] = fechamento_sem_cancelados['PRECO VENDA'] if 'PRECO VENDA' in fechamento_sem_cancelados.columns else 0

# CORREÇÃO: Preencher Quinzena usando o dicionário criado
base_df['PK'] = base_df['OS'].astype(str) + "_" + base_df['NF-E'].astype(str) + "_" + base_df['CODPRODUTO'].astype(str)
base_df['Quinzena'] = base_df['PK'].map(lambda x: quinzena_dict.get(x, ""))

base_df['GRUPO'] = base_df['GRUPO'].fillna('VAREJO')

# 1. QTDE AJUSTADA - CORREÇÃO: implementar a lógica exata do Excel
def calcular_qtde_ajustada(row):
    try:
        if row['QTDE REAL'] <= 0:
            return row['QTDE REAL']
        
        codproduto = int(row['CODPRODUTO']) if pd.notna(row['CODPRODUTO']) else None
        data = row['DATA']
        
        if codproduto is None or data is None:
            return row['QTDE REAL']
            
        custo_info = custos_dict.get((codproduto, data), {})
        peso = custo_info.get('PESO', 1)  # Padrão 1 se não encontrar
        
        # CORREÇÃO: Implementar a lógica exata do Excel
        if peso > 1:
            return row['QTDE'] * peso
        else:
            return row['QTDE REAL']
    except:
        return row['QTDE REAL']

base_df['QTDE AJUSTADA'] = base_df.apply(calcular_qtde_ajustada, axis=1)

# 2. QTDE REAL2 - CORREÇÃO: implementar a lógica exata do Excel
def calcular_qtde_real2(row):
    try:
        codproduto = int(row['CODPRODUTO']) if pd.notna(row['CODPRODUTO']) else None
        data = row['DATA']
        
        if codproduto is None or data is None:
            return np.nan
            
        custo_info = custos_dict.get((codproduto, data), {})
        peso = custo_info.get('PESO', 1)  # Padrão 1 se não encontrar
        
        if row['QTDE REAL'] < 0:
            return -row['QTDE AJUSTADA'] * peso
        else:
            return row['QTDE AJUSTADA'] * peso
    except:
        return np.nan

base_df['QTDE REAL2'] = base_df.apply(calcular_qtde_real2, axis=1)

# 3. CUSTO - COM VERIFICAÇÕES DETALHADAS
def buscar_custo(row):
    try:
        codproduto = int(row['CODPRODUTO']) if pd.notna(row['CODPRODUTO']) else None
        data = row['DATA']
        
        if codproduto is None or data is None:
            return np.nan
            
        # Verificar se a chave existe no dicionário
        key = (codproduto, data)
        if key in custos_dict:
            custo = custos_dict[key].get('CUSTO', 0)
            return custo if custo != 0 else np.nan
        else:
            return np.nan
    except Exception as e:
        return np.nan

base_df['CUSTO'] = base_df.apply(buscar_custo, axis=1)

# Verificar quantos valores de CUSTO foram encontrados
custos_encontrados = base_df['CUSTO'].notna().sum()
custos_faltantes = base_df['CUSTO'].isna().sum()

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
        data = row['DATA']
        
        if codproduto is None or data is None:
            return 0
            
        custo_info = custos_dict.get((codproduto, data), {})
        frete = custo_info.get('FRETE', 0)
        
        return frete
    except:
        return 0

base_df['Frete'] = base_df.apply(buscar_frete, axis=1)

# Verificar quantos valores de Frete foram encontrados
fretes_encontrados = (base_df['Frete'] > 0).sum()

# 6. Produção
def buscar_producao(row):
    try:
        codproduto = int(row['CODPRODUTO']) if pd.notna(row['CODPRODUTO']) else None
        data = row['DATA']
        
        if codproduto is None or data is None:
            return 0
            
        custo_info = custos_dict.get((codproduto, data), {})
        producao = custo_info.get('PRODUÇÃO', 0)
        
        return producao
    except:
        return 0

base_df['Produção'] = base_df.apply(buscar_producao, axis=1)

# Verificar quantos valores de Produção foram encontrados
producao_encontrados = (base_df['Produção'] > 0).sum()

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

# Calcular Aliq Icms - NOVA IMPLEMENTAÇÃO
base_df['Aliq Icms'] = base_df.apply(
    lambda row: round(row['VL ICMS'] / row['Fat. Bruto'], 2) if row['Fat. Bruto'] != 0 else 0, axis=1
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
# Substituir infinitos por None
base_df['frete'] = base_df['frete'].replace([np.inf, -np.inf], None)

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

# CORREÇÃO: Comissão por Regra usando lookup por múltiplas condições
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

# 24. PK (já criada anteriormente)

# 25. Coluna2
base_df['Coluna2'] = base_df['Comissão por Regra'] == base_df['Comissão Kg']

# 26. FRETE - LUC/PREJ
base_df['FRETE - LUC/PREJ'] = base_df['QTDE AJUSTADA'] * base_df['Frete']

# CORREÇÃO: DESC FEC usando lookup por PK
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

# Substituir NaN por string vazia para melhor visualização
base_df = base_df.fillna("")

# Verificar e exibir apenas linhas sem correspondência no dicionário de custos
linhas_sem_correspondencia = []

for _, row in base_df.iterrows():
    try:
        codproduto = int(row['CODPRODUTO']) if pd.notna(row['CODPRODUTO']) else None
        data = row['DATA']
        
        if codproduto is None or data is None:
            linhas_sem_correspondencia.append((data, codproduto))
            continue
            
        # Verificar se a chave existe no dicionário
        key = (codproduto, data)
        if key not in custos_dict:
            linhas_sem_correspondencia.append((data, codproduto))
    except:
        linhas_sem_correspondencia.append((data, codproduto))

# Limpar console novamente para mostrar apenas os números importantes
os.system('cls' if os.name == 'nt' else 'clear')

# Exibir apenas as contagens importantes
print("=== RESUMO DO PROCESSAMENTO ===")
print(f"Total de registros no fechamento: {len(fechamento)}")
print(f"Total de registros após remover cancelados: {len(fechamento_sem_cancelados)}")
print(f"Notas canceladas identificadas: {len(notas_canceladas)}")
print(f"Devoluções identificadas: {len(devolucoes_var)}")
print(f"Vendas identificadas: {len(vendas_var)}")
print(f"Custos de produtos carregados: {len(custos_produtos)}")
print(f"Custos sem data: {custos_produtos_sem_data}")
print(f"Custos sem código: {custos_produtos_sem_codigo}")
print(f"Custos encontrados no dicionário: {custos_encontrados}")
print(f"Custos não encontrados: {custos_faltantes}")
print(f"Fretes encontrados: {fretes_encontrados}")
print(f"Produções encontradas: {producao_encontrados}")
print(f"Registros sem correspondência de custos: {len(set(linhas_sem_correspondencia))}")

# Criar arquivo Excel
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

# Função serializadora atualizada para lidar com infinitos
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
        return None  # Tratar infinitos
    raise TypeError(f"Type {type(obj)} not serializable")

with open(json_path, 'w', encoding='utf-8') as f:
    json.dump(base_df.to_dict(orient='records'), f, ensure_ascii=False, indent=4, default=default_serializer)

print(f"\nArquivo Excel salvo em: {output_path}")
print(f"Arquivo JSON salvo em: {json_path}")
print("Processo concluído!")