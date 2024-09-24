from model.analise import calcular_beta, processar_dados_csv
from view.visualizacoes import exportar_para_excel, gerar_graficos

# Ações para calcular o beta
caminho_base = 'src/csv/SUZB3.csv'

# Processar os dados
dados_processados = processar_dados_csv(caminho_base)

# Índice de referência (Ibovespa)
indice_ibovespa = '^BVSP'
data_inicio = '2023-01-01'
data_fim = '2024-01-01'  # teste de período
beta, df_retorno = calcular_beta(indice_ibovespa, data_inicio, data_fim, dados_processados)
print(f'Beta: {beta}')

# Saída Excel
exportar_para_excel(df_retorno, beta, 'doc/resultados_beta.xlsx')

# Gerar gráficos
gerar_graficos(df_retorno)


