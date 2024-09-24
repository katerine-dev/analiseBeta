import pandas as pd
import numpy as np
import yfinance as yf # indíces da Ibovespa

# Faxina dos dados 

def processar_dados_csv(caminho_csv):
    # Ler o CSV
    dados = pd.read_csv(caminho_csv, delimiter=',')
    
    # Converter a coluna "Data" para datetime
    dados['Data'] = pd.to_datetime(dados['Data'], format='%d.%m.%Y')

    # Transformar colunas numéricas
    for coluna in ['Último', 'Abertura', 'Máxima', 'Mínima']:
        # Substituir 'n/d' e outros não numéricos por NaN
        dados[coluna] = dados[coluna].replace('n/d', np.nan)
        # Substituir vírgula por ponto para converter em float
        dados[coluna] = dados[coluna].str.replace(',', '.').astype(float)

    # Processar a coluna "Vol." (Volume)
    dados['Vol.'] = dados['Vol.'].replace(['n/d', ''], np.nan)  # Substituir 'n/d' por NaN
    # Remover 'M' e converter para float
    dados['Vol.'] = dados['Vol.'].str.replace('M', '').str.replace(',', '')

    # Converter para float e ajustar para milhões
    dados['Vol.'] = dados['Vol.'].astype(float) * 1e6

    # Processar a coluna "Var%"
    dados['Var%'] = dados['Var%'].replace('n/d', np.nan)
    # Remover o símbolo de porcentagem e converter para float
    dados['Var%'] = dados['Var%'].str.replace('%', '').str.replace(',', '.').astype(float)

    return dados

# --------------------------------- // ---------------------------------

# Função para calcular o beta de uma ação em relação ao mercado
def calcular_beta(indice_ibovespa, data_inicio, data_fim, dados_processados):
    # Converter as datas para o formato datetime
    data_inicio = pd.to_datetime(data_inicio)
    data_fim = pd.to_datetime(data_fim)
    
   # Baixar os dados dos índices da Ibovespa
    dados_mercado = yf.download(indice_ibovespa, start=data_inicio, end=data_fim)['Adj Close']  # Yahoo Finance
    
    # Calcular os retornos diários
    dados_processados.set_index('Data', inplace=True)  # Definir 'DATA' como índice
    # remove os valores faltantes, garantindo apenas os retornos válidos.
    retornos_acao = dados_processados['Último'].pct_change().dropna() # Usei a coluna 'Último' para os calculos dos retornos diários
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