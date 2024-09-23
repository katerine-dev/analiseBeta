import pandas as pd
import numpy as np
from datetime import datetime
import yfinance as yf # indíces da Ibovespa
from datetime import datetime
from pandas.tseries.offsets import DateOffset
import matplotlib.pyplot as plt
import os

# Faxina dos dados 

def processar_dados_csv(caminho_csv):
    # Ler o CSV
    dados = pd.read_csv(caminho_csv, delimiter=',')
    
    # Converter a coluna "DATA" para datetime
    dados['DATA'] = pd.to_datetime(dados['DATA'], format='%d/%m/%Y')

    # Transformar colunas numéricas
    for coluna in ['ABERTURA', 'FECHAMENTO', 'VARIAÇÃO', 'MÍNIMO', 'MÁXIMO']:
        # Substituir 'n/d' e outros não numéricos por NaN
        dados[coluna] = dados[coluna].replace('n/d', np.nan)
        dados[coluna] = dados[coluna].str.replace(',', '.').astype(float)

    # Processar a coluna "VOLUME"
    dados['VOLUME'] = dados['VOLUME'].replace(['n/d', ''], np.nan)  # Substituir 'n/d' por NaN
    # Remover 'B' e 'M' e converter para float
    dados['VOLUME'] = dados['VOLUME'].str.replace('B', '').str.replace('M', '').str.replace(',', '')

    # Converter para float e ajustar para bilhões ou milhões
    dados['VOLUME'] = dados['VOLUME'].astype(float).where(dados['VOLUME'].str.contains('B', na=False), 
                                                          dados['VOLUME'].astype(float) * 1e6)


    return dados


# Ações para calcular o beta
caminho_base = 'src/database/ITUB4.csv'

# Processar os dados
dados_processados = processar_dados_csv(caminho_base)

# --------------------------------- // ---------------------------------

# Função para calcular o beta de uma ação em relação ao mercado
def calcular_beta(indice_ibovespa, data_inicio, data_fim, dados_processados):
    # Converter as datas para o formato datetime
    data_inicio = pd.to_datetime(data_inicio)
    data_fim = pd.to_datetime(data_fim)
    
   # Baixar os dados dos índices da Ibovespa
    dados_mercado = yf.download(indice_ibovespa, start=data_inicio, end=data_fim)['Adj Close']  # Yahoo Finance
    
    # Calcular os retornos diários
    dados_processados.set_index('DATA', inplace=True)  # Definir 'DATA' como índice
    # remove os valores faltantes, garantindo apenas os retornos válidos.
    retornos_acao = dados_processados['FECHAMENTO'].pct_change().dropna() # Usei a coluna 'FECHAMENTO' para os calculos dos retornos diários
    retornos_mercado = dados_mercado.pct_change().dropna() 
    
    # Imprimir as datas dos retornos
    print("Datas dos Retornos da Ação:")
    print(retornos_acao.index)
    
    print("Datas dos Retornos do Mercado:")
    print(retornos_mercado.index)

    # Combinar os dados de retornos em um DataFrame
    df_retorno = pd.concat([retornos_acao, retornos_mercado], axis=1, join='inner').dropna()
    df_retorno.columns = ['acao', 'mercado']

    # Verificação dos retornos:
    
    print("Dados de Retornos Combinados:")
    print(df_retorno.head())
    print("Valores Nulos:")
    print(df_retorno.isnull().sum())
    
    # Calcular a variância do mercado
    var_mercado = np.var(df_retorno['mercado'])
    print(f'Variância do Mercado: {var_mercado}')
    
    if var_mercado == 0 or df_retorno.isnull().values.any():
        raise ValueError("A variância do mercado é zero ou existem valores nulos, o que impede o cálculo do beta.")
    
    # Calcular a covariância e o beta
    cov = np.cov(df_retorno['acao'], df_retorno['mercado'])[0, 1]
    beta = cov / var_mercado
    
    print("Calculando o Beta...")
    print("Retornos da Ação:")
    print(retornos_acao.head())
    print("Retornos do Mercado:")
    print(retornos_mercado.head())


    return beta, df_retorno

# --------------------------------- // ---------------------------------
def exportar_para_excel(df_retorno, beta, caminho_arquivo):
    """
    Exporta os retornos e o beta calculado para um arquivo Excel.
    
    Parâmetros:
    - df_retorno: DataFrame com os retornos da ação e do mercado
    - beta: Valor calculado do beta
    - caminho_arquivo: Caminho onde o arquivo Excel será salvo
    """
    # Criar um DataFrame para o Beta
    df_beta = pd.DataFrame({'Beta': [beta]})
    
    # Escrever os dados para o arquivo Excel
    with pd.ExcelWriter(caminho_arquivo) as writer:
        df_retorno.to_excel(writer, sheet_name='Retornos')
        df_beta.to_excel(writer, sheet_name='Beta')
    
    print(f"Dados exportados com sucesso para {caminho_arquivo}")

# --------------------------------- // ---------------------------------

def gerar_graficos(df_retorno):
    # Gerar um timestamp para garantir nomes únicos de arquivos
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    
    # Criar o gráfico de linhas
    plt.figure(figsize=(10, 6))
    
    # Plotar os retornos da ação e do mercado
    plt.plot(df_retorno.index, df_retorno['acao'], label='Retornos da Ação', color='blue')
    plt.plot(df_retorno.index, df_retorno['mercado'], label='Retornos do Mercado', color='green')
    
    # Adicionar título e legendas
    plt.title('Retornos Diários da Ação e do Mercado')
    plt.xlabel('Data')
    plt.ylabel('Retornos (%)')
    plt.legend()
    
    # Melhorar layout
    plt.grid(True)
    plt.tight_layout()
    
    # Salvar o gráfico na pasta 'doc' com um nome único
    caminho_arquivo = os.path.join('doc', f'grafico_retorno_{timestamp}.png')
    plt.savefig(caminho_arquivo)

    # Mostrar o gráfico
    plt.show()
    
    print(f"Gráfico salvo em: {caminho_arquivo}")

# --------------------------------- // ---------------------------------

# Índice de referência (Ibovespa)
indice_ibovespa = '^BVSP'
data_inicio = '2023-01-01'
data_fim = '2024-01-01'  # teste de período
beta, df_retorno = calcular_beta(indice_ibovespa, data_inicio, data_fim, dados_processados)
print(f'Beta: {beta}')


# Saída Excel

#exportar_para_excel(df_retorno, beta, 'doc/resultados_beta.xlsx')

# Gerar gráficos e exportar para Excel

# Gerar gráficos
gerar_graficos(df_retorno)