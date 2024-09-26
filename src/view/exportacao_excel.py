import pandas as pd

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
        # Obter o workbook e as planilhas
        workbook = writer.book
        worksheet_retorno = writer.sheets['Retornos']
        worksheet_beta = writer.sheets['Beta']
        
        # Ajustar as colunas para ambas as planilhas 
        for worksheet, df in zip([worksheet_retorno, worksheet_beta], [df_retorno, df_beta]):
            worksheet.set_column(0, 0, 25)  
            worksheet.set_column(1, 1, 25)
            worksheet.set_column(2, 2, 25)

             
# Função atualizada para exportar os dados e o gráfico para o Excel
def exportar_para_excel_com_grafico(df_retorno, beta, caminho_arquivo):
    """
    Exporta os retornos e o beta calculado para um arquivo Excel, juntamente com gráficos automáticos no Excel.
    
    Parâmetros:
    - df_retorno: DataFrame com os retornos da ação e do mercado
    - beta: Valor calculado do beta
    - caminho_arquivo: Caminho onde o arquivo Excel será salvo
    """
    # Converter retornos para porcentagem
    df_retorno['acao'] = df_retorno['acao'] 
    df_retorno['mercado'] = df_retorno['mercado'] 

    # Formatar a coluna de datas como "mes/ano"
    df_retorno['Data Formatada'] = df_retorno.index.to_series().dt.strftime('%m/%Y')

    # Criar um DataFrame para o Beta
    df_beta = pd.DataFrame({'Beta': [beta]})

    # Usar o XlsxWriter como engine para criar o arquivo Excel
    with pd.ExcelWriter(caminho_arquivo, engine='xlsxwriter') as writer:
        # Exportar os dados dos retornos para o Excel
        df_retorno.to_excel(writer, sheet_name='Retornos')
        df_beta.to_excel(writer, sheet_name='Beta')

        # Acessar o workbook e a worksheet do XlsxWriter
        workbook = writer.book
        worksheet = writer.sheets['Retornos']

        # Criar um gráfico de linha no Excel
        chart = workbook.add_chart({'type': 'line'})
        
        chart.add_series({
            'name': 'Retornos da Ação',
            'categories': '=Retornos!$D$2:$D${}'.format(len(df_retorno) + 1),  # Coluna de datas formatadas
            'values': '=Retornos!$B$2:$B${}'.format(len(df_retorno) + 1),      # Coluna de retornos da ação
            'line': {'width': 0.8} 
        })
        
        chart.add_series({
            'name': 'Retornos do Mercado',
            'categories': '=Retornos!$D$2:$D${}'.format(len(df_retorno) + 1),  # Coluna de datas formatadas
            'values': '=Retornos!$C$2:$C${}'.format(len(df_retorno) + 1),      # Coluna de retornos do mercado
            'line': {'width': 0.8} 
        })

        # Adicionar o título do gráfico e os eixos
        chart.set_title({'name': 'Retornos Diários da Ação e do Mercado'})
        chart.set_x_axis({
            'name': 'Data',
            'date_axis': True,            
            'num_format': 'mm/yyyy',      
            'label_position': 'low',
        })
        chart.set_y_axis({
            'name': 'Retorno (%)',
            'num_format': '0.00',  # Formato de número no eixo Y como percentual
        })

        # Ajustar o layout do gráfico para melhorar a legibilidade
        chart.set_legend({'position': 'bottom'})
        chart.set_size({'width': 720, 'height': 480})  # Ajustar o tamanho do gráfico

        # Inserir o gráfico na planilha 'Retornos'
        worksheet.insert_chart('E2', chart)
