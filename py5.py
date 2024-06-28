import pandas as pd

#code teste para apenas retorno do portfolio 1

# Caminho para o arquivo Excel
caminho_arquivo = r'C:\Users\rober\OneDrive\Documentos\Python_Technical_Test_1.xlsx'  # Substitua pelo seu caminho

# Nome da planilha
nome_da_planilha = 'Plot3'  # Substitua pelo nome da sua planilha

# Ler o arquivo Excel em um DataFrame
df = pd.read_excel(caminho_arquivo, sheet_name=nome_da_planilha, header=None)

# Selecionar a coluna do portfólio: coluna 23 (indexada como 22 no pandas)
portfolio = df.iloc[1:2195, 22]

# Converter os valores da coluna para numéricos (coercir erros para NaN)
portfolio = pd.to_numeric(portfolio, errors='coerce')

# Calcular os retornos diários começando a partir da linha 2
retornos_diarios = portfolio.pct_change().iloc[1:]

# Adicionar NaN na primeira linha para alinhar com o DataFrame original
retornos_diarios = pd.Series([None] + list(retornos_diarios), index=df.index[1:2195])

# Adicionar a coluna de retornos diários à coluna 24 (indexada como 23 no pandas)
df.iloc[1:2195, 23] = retornos_diarios

# Mostrar os retornos diários
print(retornos_diarios)

# Salvar os resultados em um novo arquivo Excel (opcional)
novo_caminho_arquivo = r'C:\Users\rober\OneDrive\Documentos\Retornos_Diarios.xlsx'  # Defina o caminho para o novo arquivo Excel
df.to_excel(novo_caminho_arquivo, index=False, header=False)

print(f"Retornos diários salvos em {novo_caminho_arquivo}")