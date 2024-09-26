VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub calcularBeta()
    Dim nIbov As Long
    Dim n As Long
    Dim retAcao() As Double
    Dim retIbov() As Double
    Dim somaRetAcao As Double
    Dim somaRetIbov As Double
    Dim somaCov As Double
    Dim somaVarIbov As Double
    Dim beta As Double
    Dim mediaRetAcao As Double
    Dim mediaRetIbov As Double
    
    ' Define a planilha com os dados das a��es (ITUB4) e da Ibovespa
    Set wsAcao = ThisWorkbook.Sheets("ITUB4")
    Set wsIbov = ThisWorkbook.Sheets("Ibovespa")
    
    ' Conta o n�mero de linhas nas tabelas
    nAcao = wsAcao.Cells(wsAcao.Rows.Count, 1).End(xlUp).Row
    nIbov = wsIbov.Cells(wsIbov.Rows.Count, 1).End(xlUp).Row
    
    ' Define o n�mero de linhas a ser processado (o menor entre nAcao e nIbov)
    n = Application.Min(nAcao, nIbov)
    
    ' Inicializa arrays para armazenar os retornos
    ReDim retAcao(1 To n - 1)
    ReDim retIbov(1 To n - 1)
    
    ' Obt�m os retornos das a��es (coluna Var%)
    For i = 2 To n
        retAcao(i - 1) = wsAcao.Cells(i, 7).Value
    Next i
    
    ' Obt�m os retornos do �ndice Ibovespa (coluna Var%)
    For i = 2 To n
        retIbov(i - 1) = wsIbov.Cells(i, 7).Value
    Next i
    
    ' Inicializar as vari�veis de soma para c�lculo
    somaRetAcao = 0
    somaRetIbov = 0
    somaCov = 0
    somaVarIbov = 0
    
    ' Calcula as somas dos retornos (para m�dias)
    For i = 1 To n - 1
        somaRetAcao = somaRetAcao + retAcao(i)
        somaRetIbov = somaRetIbov + retIbov(i)
    Next i
    
    ' Calcula as m�dias dos retornos
    mediaRetAcao = somaRetAcao / (n - 1)
    mediaRetIbov = somaRetIbov / (n - 1)
    
    ' Calcula a covari�ncia e a vari�ncia
    For i = 1 To n - 1
        somaCov = somaCov + (retAcao(i) - mediaRetAcao) * (retIbov(i) - mediaRetIbov)
        somaVarIbov = somaVarIbov + (retIbov(i) - mediaRetIbov) ^ 2
    Next i
    
    ' Calcula o Beta
    beta = somaCov / somaVarIbov
    
    ' Verifica se a aba "Beta" j� existe, caso contr�rio, cria uma nova
    On Error Resume Next
    Set wsBeta = ThisWorkbook.Sheets("Beta")
    If wsBeta Is Nothing Then
        Set wsBeta = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsBeta.Name = "Beta"
    End If
    On Error GoTo 0
    
    ' Limpa os dados existentes na aba "Beta"
    wsBeta.Cells.Clear
    
    ' Escreve o Beta calculado na aba "Beta"
    wsBeta.Cells(1, 1).Value = "Beta Calculado"
    wsBeta.Cells(2, 1).Value = beta
    
    ' Verifica se a aba "Retornos" j� existe, caso contr�rio, cria uma nova
    On Error Resume Next
    Set wsRetornos = ThisWorkbook.Sheets("Retornos")
    If wsRetornos Is Nothing Then
        Set wsRetornos = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsRetornos.Name = "Retornos"
    End If
    On Error GoTo 0
    
    ' Limpa os dados existentes na aba "Retornos"
    wsRetornos.Cells.Clear
    
    ' Escreve os retornos na aba "Retornos"
    wsRetornos.Cells(1, 1).Value = "Data"
    wsRetornos.Cells(1, 2).Value = "Retorno A��o (ITUB4)"
    wsRetornos.Cells(1, 3).Value = "Retorno Ibovespa"
    
    ' Preenche os retornos calculados nas colunas da aba "Retornos"
    For i = 2 To n
        wsRetornos.Cells(i, 1).Value = wsAcao.Cells(i, 1).Value ' Data
        wsRetornos.Cells(i, 2).Value = retAcao(i - 1) ' Retorno da A��o
        wsRetornos.Cells(i, 3).Value = retIbov(i - 1) ' Retorno do Ibovespa
    Next i
    
    ' Exibe uma mensagem de sucesso
    MsgBox "O Beta foi calculado e inserido na aba 'Beta'. Retornos foram armazenados na aba 'Retornos'."
End Sub



