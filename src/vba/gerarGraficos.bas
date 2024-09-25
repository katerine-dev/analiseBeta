Attribute VB_Name = "Module1"
Sub GerarGraficoPorPeriodo(dataInicio As Date, dataFim As Date)
    Dim wsRetornos As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim linhaInicio As Long
    Dim linhaFim As Long
    Dim chartObj As ChartObject
    Dim chart As chart
    Dim dataColuna As Range
    Dim categoriaAcao As Range
    Dim categoriaIbov As Range

    ' Definir a aba onde estão os dados
    Set wsRetornos = ThisWorkbook.Sheets("Retornos")
    
    ' Encontrar a última linha de dados na planilha "Retornos"
    
    ultimaLinha = wsRetornos.UsedRange.Rows.Count + wsRetornos.UsedRange.Row - 1

    ' Inicializar as linhas de início e fim
    linhaInicio = 0
    linhaFim = 0

    ' Encontrar a linha de início
    For i = 2 To ultimaLinha
        If wsRetornos.Cells(i, 1).Value <= dataInicio Then
            linhaInicio = i
            Exit For
        End If
    Next i

    ' Encontrar a linha de fim
    For i = linhaInicio To ultimaLinha
        If wsRetornos.Cells(i, 1).Value >= dataFim Then
            linhaFim = i - 1
            Exit For
        End If
    Next i

    ' Se linhaFim não foi encontrada, usar a última linha
    If linhaFim = 0 Then linhaFim = ultimaLinha

    ' Verificar se linhaInicio e linhaFim foram encontrados
    If linhaInicio = 0 Or linhaFim = 0 Then
        MsgBox "Nenhum dado encontrado para o período especificado."
        Exit Sub
    End If

    ' Definir as categorias (datas) e os valores (retornos) para o gráfico
    Set dataColuna = wsRetornos.Range(wsRetornos.Cells(linhaInicio, 1), wsRetornos.Cells(linhaFim, 1)) ' Coluna de datas
    Set categoriaAcao = wsRetornos.Range(wsRetornos.Cells(linhaInicio, 2), wsRetornos.Cells(linhaFim, 2)) ' Retornos ação
    Set categoriaIbov = wsRetornos.Range(wsRetornos.Cells(linhaInicio, 3), wsRetornos.Cells(linhaFim, 3)) ' Retornos Ibov
    
    ' Criar gráfico
    Set chartObj = wsRetornos.ChartObjects.Add(Left:=300, Width:=600, Top:=20, Height:=300)
    Set chart = chartObj.chart
    
    ' Definir o tipo do gráfico como linha
    chart.ChartType = xlLine
    
    ' Adicionar séries de dados
    chart.SeriesCollection.NewSeries
    chart.SeriesCollection(1).XValues = dataColuna
    chart.SeriesCollection(1).Values = categoriaAcao
    chart.SeriesCollection(1).Name = "Retornos da Ação (ITUB4)"
    
    chart.SeriesCollection.NewSeries
    chart.SeriesCollection(2).XValues = dataColuna
    chart.SeriesCollection(2).Values = categoriaIbov
    chart.SeriesCollection(2).Name = "Retornos do Ibovespa"
    
    ' Modificar a largura da linha da série de ações
    chart.SeriesCollection(1).Format.Line.Weight = 0.8 ' Aumente o número para deixar mais grossa

    ' Modificar a largura da linha da série do Ibovespa
    chart.SeriesCollection(2).Format.Line.Weight = 0.8 ' Aumente o número para deixar mais grossa
    
    ' Configurar o título e eixos do gráfico
    chart.HasTitle = True
    chart.ChartTitle.Text = "Retornos da Ação e do Ibovespa"
    chart.Axes(xlCategory).CategoryNames = dataColuna
    chart.Axes(xlCategory).HasTitle = True
    chart.Axes(xlCategory).AxisTitle.Text = "Data"
    chart.Axes(xlCategory).TickLabels.NumberFormat = "mm/yyyy"
    chart.Axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow
    
    chart.Axes(xlValue).HasTitle = True
    chart.Axes(xlValue).AxisTitle.Text = "Retorno (%)"

    ' Ajustar o tamanho do gráfico
    chartObj.Width = 700
    chartObj.Height = 400

    ' Exibir mensagem
    MsgBox "Gráfico gerado para o período de " & Format(dataInicio, "dd/mm/yyyy") & " até " & Format(dataFim, "dd/mm/yyyy")
End Sub

Sub GerarGraficoPeriodoEspecifico1anos()
    Dim dataInicio As Date
    Dim dataFim As Date

    ' Definir a data de início e fim
    dataInicio = DateValue("30/08/23")
    dataFim = DateValue("30/08/24")
    
    ' Chamar a função para gerar o gráfico
    Call GerarGraficoPorPeriodo(dataInicio, dataFim)
End Sub
Sub GerarGraficoPeriodoEspecifico3anos()
    Dim dataInicio As Date
    Dim dataFim As Date

    ' Definir a data de início e fim
    dataInicio = DateValue("29/03/21")
    dataFim = DateValue("29/03/24")
    
    ' Chamar a função para gerar o gráfico
    Call GerarGraficoPorPeriodo(dataInicio, dataFim)
End Sub
Sub GerarGraficoPeriodoEspecifico5anos()
    Dim dataInicio As Date
    Dim dataFim As Date

    ' Definir a data de início e fim
    dataInicio = DateValue("30/05/19")
    dataFim = DateValue("30/05/24")
    
    ' Chamar a função para gerar o gráfico
    Call GerarGraficoPorPeriodo(dataInicio, dataFim)
End Sub


