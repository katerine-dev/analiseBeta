from model.analise import buscar_dados_yahoo, calcular_beta, processar_dados_csv
from view.exportacao_excel import exportar_para_excel, exportar_para_excel_com_grafico
from view.graficos import gerar_graficos
# Função para especificar: função, pode especificar diferentes bases e os gráficos e arquivos Excel gerados com 
# os nomes corretos, refletindo qual base de dados foi utilizada.
def calcular_e_gerar_graficos(caminho_base, nome_base, duracao_anos, data_inicio, data_fim, nome_arquivo_excel):
    # Processar os dados da base específica
    dados_processados = processar_dados_csv(caminho_base)
    
    indice_ibovespa = '^BVSP'
    dados_mercado = buscar_dados_yahoo(indice_ibovespa, data_inicio, data_fim)
    beta, df_retorno = calcular_beta(dados_mercado, dados_processados)
    
    print(f'Beta para {duracao_anos} anos: {beta}')
    
    # Saída Excel
    exportar_para_excel(df_retorno, beta, nome_arquivo_excel)
    
    # Gerar gráficos com o nome correto da base
    gerar_graficos(df_retorno, f'{nome_base}_{duracao_anos}anos')
    
    caminho_arquivo_excel_grafico = f"doc/{nome_base}cg_{duracao_anos}anos.xlsx"
    exportar_para_excel_com_grafico(df_retorno, beta, caminho_arquivo_excel_grafico)
