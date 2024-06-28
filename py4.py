import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

#code para o heatmap

# Caminho para o arquivo Excel
caminho_arquivo = r'C:\Users\rober\OneDrive\Documentos\Python_Technical_Test_1.xlsx'  # Substitua pelo seu caminho

# Nome da planilha
nome_da_planilha = 'Plot2'  # Substitua pelo nome da sua planilha

# Ler o arquivo Excel em um DataFrame
df = pd.read_excel(caminho_arquivo, sheet_name=nome_da_planilha, header=None)

# Selecionar apenas as colunas 2 a 6 (pandas usa indexação 0, então subtrair 1)
dados = df.iloc[1:2195, 1:6]

# Definir os nomes das colunas para os gráficos e tabela
nomes_colunas = ['IVV US', 'IWM US', 'AGG US', 'AOR US', 'IAU US']
dados.columns = nomes_colunas

# Calcular a matriz de correlação
matriz_correlacao = dados.corr()

# Plotar o heatmap da matriz de correlação
plt.figure(figsize=(10, 8))
sns.heatmap(matriz_correlacao, annot=True, cmap='coolwarm', vmin=-1, vmax=1, square=True, cbar_kws={"shrink": 0.8})
plt.title('Heatmap da Matriz de Correlação')
plt.show()