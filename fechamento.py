import pandas as pd
import numpy as np
import locale

# Configurar locale para português Brasil (usar vírgula como decimal)
try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
except:
    try:
        locale.setlocale(locale.LC_ALL, 'Portuguese_Brazil.1252')
    except:
        print("Não foi possível configurar locale, usando formatação manual")

# Carregar os arquivos
fechamento_path = r"C:\Users\win11\Downloads\fechamento.csv"
movimentacao_path = r"C:\Users\win11\Downloads\movimentação.csv"

# Ler os arquivos CSV corrigindo o separador
print("Lendo arquivos...")

# Primeiro, vamos examinar a estrutura real dos arquivos
try:
    # Tentar ler com diferentes separadores
    fechamento_df = pd.read_csv(fechamento_path, encoding='utf-8', sep=';', decimal=',')
    print("Fechamento carregado com separador ; e decimal ,")
except:
    try:
        fechamento_df = pd.read_csv(fechamento_path, encoding='utf-8', sep=',', decimal=';')
        print("Fechamento carregado com separador , e decimal ;")
    except:
        try:
            # Tentar ler sem especificar decimal
            fechamento_df = pd.read_csv(fechamento_path, encoding='utf-8', sep=';')
            print("Fechamento carregado com separador ; (decimal automático)")
        except:
            # Se não funcionar, ler como string e fazer parsing manual
            with open(fechamento_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            
            # Extrair cabeçalho e dados
            header = lines[0].strip().split(';')
            data = []
            for line in lines[1:]:
                if line.strip():
                    # Substituir ponto por vírgula nos números decimais
                    cleaned_line = line.strip().replace('.', ',')
                    data.append(cleaned_line.split(';'))
            
            fechamento_df = pd.DataFrame(data, columns=header)
            print("Fechamento carregado com parsing manual e conversão de decimal")

try:
    movimentacao_df = pd.read_csv(movimentacao_path, encoding='utf-8', sep=';', decimal=',')
    print("Movimentação carregado com separador ; e decimal ,")
except:
    try:
        movimentacao_df = pd.read_csv(movimentacao_path, encoding='utf-8', sep=',', decimal=';')
        print("Movimentação carregado com separador , e decimal ;")
    except:
        try:
            movimentacao_df = pd.read_csv(movimentacao_path, encoding='utf-8', sep=';')
            print("Movimentação carregado com separador ; (decimal automático)")
        except:
            # Se não funcionar, ler como string e fazer parsing manual
            with open(movimentacao_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            
            # Extrair cabeçalho e dados
            header = lines[0].strip().split(';')
            data = []
            for line in lines[1:]:
                if line.strip():
                    # Substituir ponto por vírgula nos números decimais
                    cleaned_line = line.strip().replace('.', ',')
                    data.append(cleaned_line.split(';'))
            
            movimentacao_df = pd.DataFrame(data, columns=header)
            print("Movimentação carregado com parsing manual e conversão de decimal")

# Normalizar nomes das colunas (remover espaços extras)
fechamento_df.columns = [col.strip() for col in fechamento_df.columns]
movimentacao_df.columns = [col.strip() for col in movimentacao_df.columns]

print(f"\nColunas no fechamento.csv: {list(fechamento_df.columns)}")

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
print(fechamento_df['GRUPO'].unique())

# Verificar alguns exemplos de DESCONTO para debug
print("\nAmostra dos valores de DESCONTO (originais):")
print(fechamento_df[['GRUPO', 'DESCONTO']].head(10))

# Adicionar coluna Desconto verificado CORRIGIDA
print("\nAdicionando coluna 'Desconto verificado'...")
fechamento_df['Desconto verificado'] = fechamento_df.apply(
    lambda row: calcular_desconto_verificado(row.get('GRUPO', ''), row.get('DESCONTO', 0)), 
    axis=1
)

# Adicionar coluna Mov (comparação apenas por NOTA FISCAL/NF-E)
print("Adicionando coluna 'Mov'...")

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

# Adicionar coluna Mov V2 (comparação tripla)
print("Adicionando coluna 'Mov V2'...")

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
fechamento_df.to_csv(output_path, index=False, encoding='utf-8', sep=';', decimal=',')

print(f"\nProcessamento concluído!")
print(f"Arquivo salvo em: {output_path}")
print(f"Colunas adicionadas: Desconto verificado, Mov, Mov V2")
print(f"Total de linhas processadas: {len(fechamento_df)}")

# Mostrar amostra dos resultados com foco no Desconto verificado
print(f"\nAmostra dos resultados - Desconto verificado:")
print(fechamento_df[['GRUPO', 'DESCONTO', 'Desconto verificado']].head(15))

# Mostrar estatísticas do Desconto verificado
print(f"\nEstatísticas do Desconto verificado:")
print(f"Valores únicos: {fechamento_df['Desconto verificado'].unique()}")