import pandas as pd
from datetime import datetime
from matplotlib.dates import DateFormatter
import matplotlib.pyplot as plt
import os

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



def gerar_graficos(df_retorno):
    # Gerar um timestamp para garantir nomes únicos de arquivos
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    
    # Criar o gráfico de linhas
    plt.figure(figsize=(10, 6))
    
    # Plotar os retornos da ação e do mercado
    plt.plot(df_retorno.index, df_retorno['acao'], label='Retornos da Ação', color='#042453', linewidth=0.5)
    plt.plot(df_retorno.index, df_retorno['mercado'], label='Retornos do Mercado', color='grey', alpha=0.7, linewidth=0.5)
    

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