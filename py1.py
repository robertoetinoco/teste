import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

# code para plotar gráfico de evolução da série histórica de preços

# Carregar o arquivo Excel
caminho_arquivo = r'C:\Users\rober\OneDrive\Documentos\Python_Technical_Test_1.xlsx'  # Use string bruta para evitar problemas com caracteres de escape
nome_da_planilha = 'Plot'  # Substitua pelo nome da sua planilha, se necessário

# Ler o arquivo Excel em um DataFrame
df = pd.read_excel(caminho_arquivo, sheet_name=nome_da_planilha, header=None)

# Selecionar o intervalo de dados: linhas 2 a 2195 e colunas 2 a 6 (pandas usa indexação 0, então subtrair 1)
dados = df.iloc[1:2195, 1:6]

# Definir nomes das colunas para melhor legibilidade (opcional)
dados.columns = ['IVV US', 'IWM US', 'AGG US', 'AOR US', 'IAU US']

# Criar uma faixa de datas para o eixo x (supondo que cada linha represente um dia consecutivo)
datas = pd.date_range(start='2017-12-29', periods=len(dados), freq='D')  # Ajuste a data de início conforme necessário

# Plotar os dados
plt.figure(figsize=(12, 8))

for coluna in dados.columns:
    plt.plot(datas, dados[coluna], label=coluna)

# Adicionar título e rótulos
plt.title('Evolução dos Preços dos Ativos por Trimestre')
plt.xlabel('Data')
plt.ylabel('Preço')
plt.ylim(20, 600)  # Ajustar o eixo y para o intervalo de 20 a 600
plt.legend()

# Ajustar o formato do eixo x para mostrar apenas trimestres de forma mais discreta
locator = mdates.MonthLocator(bymonth=[1, 4, 7, 10], bymonthday=1)  # Mostrar apenas no primeiro dia do trimestre
plt.gca().xaxis.set_major_locator(locator)
plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%b %Y'))

# Rotacionar os rótulos do eixo x para melhor legibilidade
plt.xticks(rotation=45)

# Mostrar o gráfico
plt.tight_layout()
plt.show()