import pandas as pd
import numpy as np
import yfinance as yf # indíces da Ibovespa

def processar_dados_csv(caminho_arquivo):
    """
    Função para faxinar os dados retirando informação n/d e transformando colunas em float
    Parâmetro: caminho_csv, caminho de diretório para o csv
    """
    # Verificar a extensão do arquivo
    if caminho_arquivo.endswith('.csv'):
        # Ler o CSV
        dados = pd.read_csv(caminho_arquivo, delimiter=',')
    elif caminho_arquivo.endswith('.xlsx') or caminho_arquivo.endswith('.xls'):
        # Ler o Excel
        dados = pd.read_excel(caminho_arquivo)
    else:
        raise ValueError("Formato de arquivo não suportado. Use um arquivo CSV ou Excel.")
    
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


def buscar_dados_yahoo(ticker, data_inicio, data_fim):
    """
    Função para buscar dados de mercado no Yahoo Finance.
    
    Parâmetros:
    - ticker: Código do ativo ou índice no Yahoo Finance (ex: '^BVSP' para o Ibovespa)
    - data_inicio: Data inicial para busca dos dados
    - data_fim: Data final para busca dos dados
    
    Retorna:
    - Série de preços ajustados de fechamento ('Adj Close') do ativo ou índice.
    """
    # Converter as datas para o formato datetime
    data_inicio = pd.to_datetime(data_inicio)
    data_fim = pd.to_datetime(data_fim)
    
    # Baixar os dados de fechamento ajustado ('Adj Close')
    dados_mercado = yf.download(ticker, start=data_inicio, end=data_fim)['Adj Close']
    
    return dados_mercado


def calcular_beta(dados_mercado, dados_processados):
    """
    Função para calcular o beta de uma ação em relação ao mercado
    """
    # Calcular os retornos diários
    dados_processados.set_index('Data', inplace=True)  # Definir 'Data' como índice
    # remove os valores faltantes, garantindo apenas os retornos válidos.
    retornos_acao = dados_processados['Último'].pct_change().dropna() # Usei a coluna 'Último' para os calculos dos retornos diários
    retornos_mercado = dados_mercado.pct_change().dropna() 
    
    # Combinar os dados de retornos em um DataFrame
    df_retorno = pd.concat([retornos_acao, retornos_mercado], axis=1, join='inner').dropna()
    df_retorno.columns = ['acao', 'mercado']

    # Calcular a variância do mercado
    var_mercado = np.var(df_retorno['mercado'])
    
    if var_mercado == 0 or df_retorno.isnull().values.any():
        raise ValueError("A variância do mercado é zero ou existem valores nulos, o que impede o cálculo do beta.")
    
    # Calcular a covariância e o beta
    cov = np.cov(df_retorno['acao'], df_retorno['mercado'])[0, 1]
    beta = cov / var_mercado

    return beta, df_retorno