from datetime import datetime
import matplotlib.pyplot as plt
import os

def gerar_graficos(df_retorno, nome_base):
  
    # Criar o gráfico de linhas
    plt.figure(figsize=(10, 6))
    
    # Plotar os retornos da ação e do mercado
    plt.plot(df_retorno.index, df_retorno['acao'], label='Retornos da Ação', color='#042453', linewidth=0.8)
    plt.plot(df_retorno.index, df_retorno['mercado'], label='Retornos do Mercado', color='#ffa500', alpha=0.7, linewidth=0.8)
    

    # Adicionar título e legendas
    plt.title('Retornos Diários da Ação e do Mercado')
    plt.xlabel('Data')
    plt.ylabel('Retornos (%)')
    plt.legend()

    # Melhorar layout
    plt.grid(True)
    plt.tight_layout()
    
    # Salvar o gráfico na pasta 'doc' com um nome único
    caminho_arquivo = os.path.join('doc', f'{nome_base}.png')
    plt.savefig(caminho_arquivo)

    # Mostrar o gráfico
    plt.show()