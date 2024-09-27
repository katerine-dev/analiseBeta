import yfinance as yf
import pandas as pd

# Definir o ticker do Ibovespa
ibovespa_ticker = '^BVSP'

# Baixar os dados históricos do Ibovespa
historico = yf.download(ibovespa_ticker, period='5y', interval='1d')

# Verificar os primeiros registros
print(historico.head())

# Definir o nome do arquivo Excel
arquivo_excel = 'ibovespa_historico.xlsx'

# Exportar os dados para o Excel
historico.to_excel(arquivo_excel)
