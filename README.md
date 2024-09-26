# Relatório - Análise Beta de Ações
Analise de cálculo do beta de 5 ações em períodos de 1, 3 e 5 anos, com base no índice Ibovespa.

## Cálculo de Beta de ações

A proposta do projeto é realizar o cálculo de Beta de 5 ações SUZB3, PETR4, ITUB4, LREN3 e WEGE3 em períodos fixos de 1, 3 e 5 anos, utilizando os índices Ibovespa.

O teste foi realizado em dois caminhos:

- **VBA**: Criar uma macro que calcule automaticamente o beta para 1, 3 e 5 anos das 5 ações, atualize os gráficos no Excel, e inclua um botão para refazer os cálculos e gráficos ao ser clicado.

- **Python**: Desenvolver um script que importe dados de CSV ou Excel, calcule os betas para 1, 3 e 5 anos, e exporte os resultados e gráficos prontos para uso em um novo arquivo Excel.

Para cada ação, serão gerados:

- Python
  - Um Excel com os resultados, sem gráfico.
  - Um Excel com os resultados e gráficos de análise.
  - Um gráfico utilizando a biblioteca Python Matplotlib.

Exemplo de gráfico:

![Gráfico - Python]('doc/LREN3_1anos.png')

- VBA
  - Um .xlms com os resultados e graficos

Exemplo de gráfico:

![Gráfico - Excel]('doc/analise_github.png')

O objetivo é comparar os formatos e a interpretação dos dados em diferentes representações. A análise do beta é uma medida de sensibilidade de um ativo em relação ao comportamento de um índice. No projeto foi utilizado os índices da Ibovespa. O objetivo é medir a volatilidade de uma ação, sendo uma das principais ferramentas para investidores que desejam entender o risco associado a cada ativo.

- Excel Sem Gráfico: Este arquivo fornecerá uma visão clara e concisa dos dados calculados, permitindo uma rápida consulta aos valores de beta ao longo dos períodos analisados.

- Excel Com Gráfico: Este arquivo incluirá gráficos que ilustram a relação entre o retorno das ações e do índice. Gráficos de linhas serão utilizados para destacar tendências e variações no beta ao longo do período analisado. Para automatização desses gráficos foi utilizado linguagem de programação Python e VBA

- Gráfico Utilizando Matplotlib: Este gráfico será gerado em Python usando a biblioteca Matplotlib e terá como objetivo representar graficamente o comportamento do beta ao longo do tempo.

Ao final da análise, será possível tirar conclusões sobre a sensibilidade das ações escolhidas em relação às variações do mercado, ajudando os investidores a tomar decisões mais informadas sobre suas carteiras de investimento. O uso de diferentes formatos de apresentação permitirá não apenas uma compreensão profunda dos dados, mas também uma comparação entre as ferramentas utilizadas, destacando suas respectivas vantagens e limitações, e contribuindo para uma escolha mais adequada das metodologias de análise em futuras avaliações de risco e desempenho de ativos.

## Dados:

- Data inicial dos dados para análise: 01/01/2019 - 23/08/2024.

- Fonte:

- Investing: [https://br.investing.com](https://br.investing.com) (SUZB3, PETR4, ITUB4, LREN, 3WEGE3, Ibovespa)

- Yahoo Finance - Finance's API: [https://pypi.org](https://pypi.org/project/yfinance/)

## Python:

### Desenvolvimento

O desenvolvimento do `analiseBeta` foi realizado utilizando a linguagem de programação Python. O projeto foi estruturado em diferentes pacotes para separar as responsabilidades e facilitar a manutenção e extensibilidade do código. (csv, model, view, vba)

Para o gerenciamento de depêndencias e ambientes foi utilizado a biblioteca [Poetry](https://python-poetry.org/) e para facilitar o controle de versões e desenvolvimento colaborativo o projeto está disponível no Github[^2].

### Decisões de funcionalidades:

O Sistema de Cálculo de Beta possui as seguintes funcionalidades principais:

1. Model:
- Permite faxinar os dados: `processar_dados_csv()`
- Buscar dados de mercado no Yahoo Finance: `buscar_dados_yahoo()`
- Calcula o Beta de uma ação: `calcular_beta()`

2. View:
- Exporta os retornos e o beta calculado para um arquivo excel: `exportar_para_excel()`
- Exporta os retornos e o beta para o excel automatizando a criação dos gráficos: `exportar_para_excel_com_grafico()`
- Cria gráficos com a biblioteca `matplotlib`: `gerar_graficos()`

3. Source: 
- Organiza todo o sistema de automatrização, facilitando a manipulação dos dados em um período determinado: `gerar_graficos_anos()`

### Uso

#### Venv

Ativando o ambiente virtual:

```sh
poetry shell
```

Adicionando novas dependências:

```sh
# O linha abaixo adiciona novas bibliotecas the `requests` library
poetry add pandas
```

Instalando dependências do projeto:
```sh
poetry install
```

Exemplo de chamada da função exportar_para_excel_com_grafico
```
from gerar_graficos_anos import calcular_e_gerar_graficos

# CÁLCULO DO BETA
# SUZB3
caminho_base_suzb3 = 'src/csv/SUZB3.csv'
calcular_e_gerar_graficos(caminho_base_suzb3, 'SUZB3', 1, '2023-05-13', '2024-05-03', 'doc/SUZB3sg_1ano.xlsx')
calcular_e_gerar_graficos(caminho_base_suzb3, 'SUZB3', 3, '2021-07-31', '2024-07-31', 'doc/SUZB3sg_3anos.xlsx')
calcular_e_gerar_graficos(caminho_base_suzb3, 'SUZB3', 5, '2019-08-28', '2024-08-28', 'doc/SUZB3sg_5anos.xlsx')

```

## VBA

Também foi utilizada uma macro em VBA para auxiliar na manipulação e formatação dos dados no Excel.

Foi necessário fazer algumas alterações manualmente:
- Conversão das células para number e percent.
- Converter data para Date  (replace nas datas para `.` para `/`)
- Inclui uma nova aba para os índices da ibovespa

Foram implementadas essas rotinas:

- Sub `calcularBeta()`
- Sub `GerarGraficoPorPeriodo()`

Para gerar automaticamente os 3 gráficos diferentes foram criadas em outro módulo:
- Sub `GerarGraficoPeriodoEspecifico1anos()`
- Sub `GerarGraficoPeriodoEspecifico3anos()`
- Sub `GerarGraficoPeriodoEspecifico4anos()`

## Calculo do BETA:

O cálculo do beta de um ativo envolve vários passos:

- Coletar dados históricos dos retornos do ativo e do índice de referência.
- Calcular os retornos do ativo e do índice.
- Calcular a covariância entre os retornos do ativo e do índice.
- Calcular a variância do índice (*Ibovespa*).

### Implementação do cálculo no Python:

A função calcular_beta(dados_mercado, dados_processados) executa os seguintes passos:

- Cálculo dos Retornos Diários:
Os dados são ajustados para definir `'Data'` como o índice do DataFrame.
Utiliza-se a função `pct_change()` para calcular os retornos diários da coluna `'Último'`, removendo valores nulos com `dropna()`. Isso garante que apenas dados válidos sejam considerados.

- Combinação dos Dados:
Os retornos da ação e do mercado são combinados em um DataFrame, permitindo cálculos de covariância e variância.

- Cálculo da Variância do Mercado:
A variância dos retornos do mercado é calculada utilizando `np.var()`. Se a variância for zero ou se houver valores nulos, um erro é levantado, garantindo a integridade dos cálculos.

- Cálculo da Covariância e do Beta:
A covariância entre os retornos da ação e do mercado é calculada usando `np.cov()`, e o Beta é obtido pela divisão da covariância pela variância do mercado.

### Implementação em VBA

- Cálculo dos Retornos Diários:
Os retornos são obtidos da coluna de variação percentual das ações e do Ibovespa, dividindo por 100 para convertê-los em formato decimal.

- Inicialização de Variáveis:
Arrays são criados para armazenar os retornos, e variáveis de soma são inicializadas para o cálculo da média.

- Cálculo das Médias:
As médias dos retornos são calculadas somando os retornos e dividindo pelo número total de períodos.
Cálculo da Covariância e Variância:
A covariância é calculada somando os produtos das diferenças dos retornos em relação às médias. A variância do Ibovespa é calculada somando os quadrados das diferenças.

- Cálculo do Beta:
O Beta é calculado ao final, e os resultados são armazenados em uma nova aba chamada "Beta".

## Estrutura do Projeto
O projeto foi estruturado em diferentes pacotes para separar as responsabilidades e facilitar a manutenção e extensibilidade do código.

- `csv`: Contém todos os arquivos .csv
- `model`: Códigos referente a modelagem do projeto
- `view`: Códigos referente todas as visualizações dos gráficos
- `vba`: Contém todas as planilhas `.xlsm` e os códigos responsáveis pelas macros.

Existe uma pasta doc (os resultados gerados com Python).

Para utilizar do sistema compile o arquivo Main.