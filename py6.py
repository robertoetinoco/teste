import pandas as pd

# Caminho para o arquivo Excel
caminho_arquivo = r'C:\Users\rober\OneDrive\Documentos\Python_Technical_Test_1.xlsx'  # Substitua pelo seu caminho

# Nome da planilha
nome_da_planilha = 'Plot3'  # Substitua pelo nome da sua planilha

# Ler o arquivo Excel em um DataFrame
df = pd.read_excel(caminho_arquivo, sheet_name=nome_da_planilha, header=None)

def processar_retornos(coluna_original, coluna_retorno, coluna_volatilidade):
    # Assegurar que as colunas de retorno e volatilidade existem
    if coluna_retorno >= len(df.columns):
        df[coluna_retorno] = None  # Adiciona uma nova coluna com valores None
    if coluna_volatilidade >= len(df.columns):
        df[coluna_volatilidade] = None  # Adiciona uma nova coluna com valores None

    # Selecionar a coluna do portfólio
    portfolio = df.iloc[1:2195, coluna_original]

    # Converter os valores da coluna para numéricos (coercir erros para NaN)
    portfolio = pd.to_numeric(portfolio, errors='coerce')

    # Calcular os retornos diários começando a partir da linha 2
    retornos_diarios = portfolio.pct_change()

    # Adicionar NaN na primeira linha para alinhar com o DataFrame original
    df.iloc[1, coluna_retorno] = None
    df.iloc[2:2195, coluna_retorno] = retornos_diarios.values[1:]

    # Calcular a volatilidade (desvio padrão anualizado dos retornos diários)
    volatilidade = retornos_diarios.std() * (252**0.5)  # Anualizando a volatilidade
    df.iloc[1:2195, coluna_volatilidade] = volatilidade

# Aplicar a função para as colunas especificadas
pares_de_colunas = [(22, 23, 24), (31, 32, 33), (40, 41, 42), (49, 50, 51)]
for orig, ret, vol in pares_de_colunas:
    processar_retornos(orig, ret, vol)

# Salvar os resultados em um novo arquivo Excel (opcional)
novo_caminho_arquivo = r'C:\Users\rober\OneDrive\Documentos\Retornos_Diarios.xlsx'  # Defina o caminho para o novo arquivo Excel
df.to_excel(novo_caminho_arquivo, index=False, header=False)

print(f"Retornos diários e volatilidade salvos em {novo_caminho_arquivo}")