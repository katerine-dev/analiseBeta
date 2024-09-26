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

    ' Encontrar a linha de início (primeira data maior ou igual a dataInicio)
    For i = 2 To ultimaLinha
        If IsDate(wsRetornos.Cells(i, 1).Value) Then ' Verifica se é uma data
            If wsRetornos.Cells(i, 1).Value >= dataInicio Then
                linhaInicio = i
                Exit For
            End If
        End If
    Next i

    ' Verifica se linhaInicio foi encontrada
    If linhaInicio = 0 Then
        MsgBox "Nenhum dado encontrado para a data de início especificada."
        Exit Sub
    End If

    ' Encontrar a linha de fim (última data menor ou igual a dataFim)
    linhaFim = linhaInicio ' Inicializa linhaFim com linhaInicio
    For i = linhaInicio To ultimaLinha
        If IsDate(wsRetornos.Cells(i, 1).Value) Then ' Verifica se é uma data
            If wsRetornos.Cells(i, 1).Value <= dataFim Then
                linhaFim = i
            Else
                Exit For
            End If
        End If
    Next i

    ' Verificar se linhaFim foi encontrada e é válida
    If linhaFim < linhaInicio Then
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
    With chart.SeriesCollection
        .NewSeries
        .Item(1).XValues = dataColuna
        .Item(1).Values = categoriaAcao
        .Item(1).Name = "Retornos da Ação "
        .NewSeries
        .Item(2).XValues = dataColuna
        .Item(2).Values = categoriaIbov
        .Item(2).Name = "Retornos do Ibovespa"
    End With
    
    ' Modificar a largura da linha das séries
    chart.SeriesCollection(1).Format.Line.Weight = 0.8 ' Ajuste conforme necessário
    chart.SeriesCollection(2).Format.Line.Weight = 0.8 ' Ajuste conforme necessário
    
    ' Configurar o título e eixos do gráfico
    With chart
        .HasTitle = True
        .ChartTitle.Text = "Retornos da Ação e do Ibovespa"
        
        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Text = "Data"
            .TickLabels.NumberFormat = "mm/yyyy"
            .TickLabelPosition = xlTickLabelPositionLow
        End With
        
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "Retorno (%)"
        End With
    End With

    ' Ajustar o tamanho do gráfico
    chartObj.Width = 700
    chartObj.Height = 400

    ' Exibir mensagem
    MsgBox "Gráfico gerado para o período de " & Format(dataInicio, "dd/mm/yyyy") & " até " & Format(dataFim, "dd/mm/yyyy")
End Sub

Sub GerarGraficoPeriodoEspecifico1anos()
    Dim dataInicio As Date
    Dim dataFim As Date

    ' Definir a data de in�cio e fim
    dataInicio = DateValue("30/08/23")
    dataFim = DateValue("30/08/24")
    
    ' Chamar a fun��o para gerar o gr�fico
    Call GerarGraficoPorPeriodo(dataInicio, dataFim)
End Sub
Sub GerarGraficoPeriodoEspecifico3anos()
    Dim dataInicio As Date
    Dim dataFim As Date

    ' Definir a data de in�cio e fim
    dataInicio = DateValue("29/03/21")
    dataFim = DateValue("29/03/24")
    
    ' Chamar a fun��o para gerar o gr�fico
    Call GerarGraficoPorPeriodo(dataInicio, dataFim)
End Sub
Sub GerarGraficoPeriodoEspecifico5anos()
    Dim dataInicio As Date
    Dim dataFim As Date

    ' Definir a data de in�cio e fim
    dataInicio = DateValue("30/05/19")
    dataFim = DateValue("30/05/24")
    
    ' Chamar a fun��o para gerar o gr�fico
    Call GerarGraficoPorPeriodo(dataInicio, dataFim)
End Sub


