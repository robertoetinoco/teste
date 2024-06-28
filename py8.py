import pandas as pd
import matplotlib.pyplot as plt

#code completo calculando retorno, vol, indice de sharpe e max drawdown - no final do dataframe

# Caminho para o arquivo Excel
caminho_arquivo = r'C:\Users\rober\OneDrive\Documentos\Python_Technical_Test_1.xlsx'

# Nome da planilha
nome_da_planilha = 'Plot3'

# Ler o arquivo Excel em um DataFrame
df = pd.read_excel(caminho_arquivo, sheet_name=nome_da_planilha, header=None)

def calcular_max_drawdown(retornos):
    cum_max = retornos.cummax()
    drawdowns = (retornos - cum_max) / cum_max
    max_drawdown = drawdowns.min()
    return max_drawdown

def processar_retornos(coluna_original, coluna_retorno, coluna_sharpe):
    # Assegurar que as colunas de retorno e de Índice Sharpe existam
    while coluna_retorno >= len(df.columns) or coluna_sharpe >= len(df.columns):
        df.insert(len(df.columns), 'new_col', None)  # Adiciona colunas novas se necessário
    
    # Selecionar a coluna do portfólio
    portfolio = df.iloc[1:2195, coluna_original]

    # Converter os valores da coluna para numéricos
    portfolio = pd.to_numeric(portfolio, errors='coerce')

    # Calcular os retornos diários
    retornos_diarios = portfolio.pct_change()

    # Atribuir os retornos diários e calcular o Índice Sharpe
    df.iloc[2:2195, coluna_retorno] = retornos_diarios.values[1:]  # Ignora o primeiro valor NaN
    media_retornos = retornos_diarios.mean()
    desvio_retornos = retornos_diarios.std()
    sharpe_anualizado = (media_retornos * 252) / (desvio_retornos * (252 ** 0.5))
    df.iloc[1, coluna_sharpe] = sharpe_anualizado  # Sharpe na linha 2

    # Calcular e armazenar o máximo drawdown
    max_drawdown = calcular_max_drawdown(retornos_diarios)
    last_row_index = df.shape[0]  # Identificar a última linha
    df.loc[last_row_index, coluna_retorno] = max_drawdown  # Armazenar drawdown na última linha

# Configurar pares de colunas
pares_de_colunas = [(22, 23, 24), (31, 32, 33), (40, 41, 42), (49, 50, 51)]
for original, retorno, sharpe in pares_de_colunas:
    processar_retornos(original, retorno, sharpe)

# Criar gráfico de barras para Índice Sharpe
sharpe_values = [df.iloc[1, sharpe] for _, _, sharpe in pares_de_colunas]
labels = [f'Coluna {sharpe}' for _, _, sharpe in pares_de_colunas]
sharpe_values, labels = zip(*sorted(zip(sharpe_values, labels), reverse=True))

plt.figure(figsize=(10, 6))
plt.bar(labels, sharpe_values, color='green')
plt.xlabel('Colunas de Índice Sharpe')
plt.ylabel('Índice Sharpe Anualizado')
plt.title('Índice Sharpe Anualizado por Coluna')
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()

# Salvar os resultados em um novo arquivo Excel
novo_caminho_arquivo = r'C:\Users\rober\OneDrive\Documentos\Retornos_Diarios.xlsx'
df.to_excel(novo_caminho_arquivo, index=False, header=False)
print(f"Índices Sharpe anualizados, retornos diários e máximos drawdowns salvos em {novo_caminho_arquivo}")