from model.analise import buscar_dados_yahoo, calcular_beta, processar_dados_csv
from view.exportacao_excel import exportar_para_excel, exportar_para_excel_com_grafico
from view.graficos import gerar_graficos

# Ações para calcular o beta
caminho_base = 'src/csv/SUZB3.csv'

# Processar os dados
dados_processados = processar_dados_csv(caminho_base)

# Índice de referência (Ibovespa)
indice_ibovespa = '^BVSP'
data_inicio = '2023-01-01'
data_fim = '2024-01-01'  # teste de período


dados_mercado = buscar_dados_yahoo(indice_ibovespa, data_inicio, data_fim)

beta, df_retorno = calcular_beta(dados_mercado, dados_processados)
print(f'Beta: {beta}')

# Saída Excel
#exportar_para_excel(df_retorno, beta, 'doc/resultados_beta.xlsx')

# Gerar gráficos
#gerar_graficos(df_retorno)
caminho_arquivo_excel = "doc/comgraficos.xlsx"

exportar_para_excel_com_grafico(df_retorno, beta, caminho_arquivo_excel)
