import pandas as pd

#code para mostrar resultados da média, mediana e desv pad - vai estar no terminal

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

# Multiplicar por 100 para expressar em percentual
dados = dados * 100

# Calcular estatísticas descritivas para cada coluna, ignorando valores zero
estatisticas = {
    nomes_colunas[0]: dados[nomes_colunas[0]][dados[nomes_colunas[0]] != 0].agg(['mean', 'median', 'std']),
    nomes_colunas[1]: dados[nomes_colunas[1]][dados[nomes_colunas[1]] != 0].agg(['mean', 'median', 'std']),
    nomes_colunas[2]: dados[nomes_colunas[2]][dados[nomes_colunas[2]] != 0].agg(['mean', 'median', 'std']),
    nomes_colunas[3]: dados[nomes_colunas[3]][dados[nomes_colunas[3]] != 0].agg(['mean', 'median', 'std']),
    nomes_colunas[4]: dados[nomes_colunas[4]][dados[nomes_colunas[4]] != 0].agg(['mean', 'median', 'std']),
}

# Criar um DataFrame com as estatísticas
estatisticas_df = pd.DataFrame(estatisticas)

# Exibir a tabela
print(estatisticas_df)