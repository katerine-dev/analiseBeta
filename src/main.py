from gerar_graficos_anos import calcular_e_gerar_graficos

"""
Cálculo do beta de 5 ações SUZB3, PETR4, ITUB4, LREN3 e WEGE3 em períodos fixos de 1, 3 e 5 anos, utilizando o índice Ibovespa
Para cada ação será gerado: 
- Um excel com os resultados, sem gráfico. 
- Um excel com os resultados e gráficos de análise.
- Um gráfico utilizando a biblioteca python matplotlib. 
O objetivo é comparar os formatos e a interpretação dos dados em diferentes representações.
"""
# CÁLCULO DO BETA
# SUZB3
# caminho_base_suzb3 = 'src/csv/SUZB3.csv'
# calcular_e_gerar_graficos(caminho_base_suzb3, 'SUZB3', 1, '2023-05-13', '2024-05-03', 'doc/SUZB3sg_1ano.xlsx')
# calcular_e_gerar_graficos(caminho_base_suzb3, 'SUZB3', 3, '2021-07-31', '2024-07-31', 'doc/SUZB3sg_3anos.xlsx')
# calcular_e_gerar_graficos(caminho_base_suzb3, 'SUZB3', 5, '2019-08-28', '2024-08-28', 'doc/SUZB3sg_5anos.xlsx')

# ITUB4
caminho_base_itub4 = 'src/csv/ITUB4.csv'
calcular_e_gerar_graficos(caminho_base_itub4, 'ITUB4', 1, '2023-08-30', '2024-08-30', 'doc/ITUB4sg_1ano.xlsx')
calcular_e_gerar_graficos(caminho_base_itub4, 'ITUB4', 3, '2021-03-29', '2024-03-29', 'doc/ITUB4sg_3anos.xlsx')
calcular_e_gerar_graficos(caminho_base_itub4, 'ITUB4', 5, '2019-05-30', '2024-05-30', 'doc/ITUB4sg_5anos.xlsx')

# # PETR4
# caminho_base_petr4 = 'src/csv/PETR4.csv'
# calcular_e_gerar_graficos(caminho_base_petr4, 'PETR4', 1, '2023-04-03', '2024-04-03', 'doc/PETR4sg_1ano.xlsx')
# calcular_e_gerar_graficos(caminho_base_petr4, 'PETR4', 3, '2021-07-16', '2024-07-16', 'doc/PETR4sg_3anos.xlsx')
# calcular_e_gerar_graficos(caminho_base_petr4, 'PETR4', 5, '2019-09-01', '2024-09-01', 'doc/PETR4sg_5anos.xlsx')

# # LREN3
# caminho_base_lren3 = 'src/csv/LREN3.csv'
# calcular_e_gerar_graficos(caminho_base_lren3, 'LREN3', 1, '2023-02-13', '2024-02-13', 'doc/LREN3sg_1ano.xlsx')
# calcular_e_gerar_graficos(caminho_base_lren3, 'LREN3', 3, '2021-01-01', '2024-01-01', 'doc/LREN3sg_3anos.xlsx')
# calcular_e_gerar_graficos(caminho_base_lren3, 'LREN3', 5, '2019-06-30', '2024-06-30', 'doc/LREN3sg_5anos.xlsx')

# # WEGE3
# caminho_base_wege3 = 'src/csv/WEGE3.csv'
# calcular_e_gerar_graficos(caminho_base_wege3, 'WEGE3', 1, '2023-05-02', '2024-05-02', 'doc/WEGE3sg_1ano.xlsx')
# calcular_e_gerar_graficos(caminho_base_wege3, 'WEGE3', 3, '2021-09-21', '2024-09-21', 'doc/WEGE3sg_3anos.xlsx')
# calcular_e_gerar_graficos(caminho_base_wege3, 'WEGE3', 5, '2019-05-07', '2024-05-07', 'doc/WEGE3sg_5anos.xlsx')
