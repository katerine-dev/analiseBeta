import pandas as pd
import numpy as np
from datetime import datetime
import yfinance as yf # indíces da Ibovespa
from datetime import datetime
from pandas.tseries.offsets import DateOffset
import matplotlib.pyplot as plt

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


    return beta

# --------------------------------- // ---------------------------------


# --------------------------------- // ---------------------------------

# Índice de referência (IBovespa)
indice_ibovespa = '^BVSP'
data_inicio = '2023-01-01'
data_fim = '2024-01-01'  # teste de período
beta = calcular_beta(indice_ibovespa, data_inicio, data_fim, dados_processados)
print(f'Beta: {beta}')



# Arquivo de saída Excel
arquivo_saida = 'doc/beta.xlsx'

# Gerar gráficos e exportar para Excel
#gerar_graficos_e_exportar_excel(acoes, mercado, arquivo_saida)
#print("Cálculo e gráficos concluídos. Resultados exportados para o arquivo Excel.")