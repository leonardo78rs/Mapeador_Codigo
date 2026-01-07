'--------------------------------------------------------------------------
' GroupFunc é uma função para agrupar e contar quantas vezes a função contém a expressão procurada
' O resultado é colocado na planilha Resumo

Sub GroupFunc(lexibemsg As Boolean)

Dim wOcorr As Worksheet
Dim wResumo As Worksheet
Dim lin1, lin2, col As Integer
Dim nini, nfim, nCountGroup As Double
Dim cFonte, cFunct As String
Dim cMensagem As String
Dim cMensagem1 As String
Dim x As Integer
Dim n As Integer
Dim nSumFunc As Integer
Dim nSumOcor As Integer
Dim nPagina As Integer

Set wOcorr = Sheets("Ocorrencias")
Set wResumo = Sheets("Resumo")
nCountGroup = 0

' Primeira linha tem que pegar separado para fazer a comparação no FOR abaixo
cFonte = wOcorr.Cells(13, 1)
cFunct = wOcorr.Cells(13, 4)
nCountGroup = 1
lin2 = 3                            ' Linha do resumo que começa a gravar

If lAnimado Then
   wOcorr.Activate
End If

For lin1 = 14 To 500

    If wOcorr.Cells(lin1, 1) = cFonte And _
       wOcorr.Cells(lin1, 4) = cFunct Then
           
       cFonte = wOcorr.Cells(lin1, 1)
       cFunct = wOcorr.Cells(lin1, 4)
       nCountGroup = nCountGroup + 1
    Else
    
       wResumo.Cells(lin2, 1) = cFonte
       wResumo.Cells(lin2, 2) = cFunct
       wResumo.Cells(lin2, 3) = nCountGroup
       lin2 = lin2 + 1
              
       cFonte = wOcorr.Cells(lin1, 1)
       cFunct = wOcorr.Cells(lin1, 4)
       nCountGroup = 1
    End If

Next

'If lAnimado Then
'   wResumo.Activate
'End If

If lexibemsg Then

    x = 3
    nSumOcor = 0
    nPagina = 0
    
    cMensagem = "Busca:" + Str(wResumo.Cells(2, 11)) + Chr(13) + Chr(13) + "Funções Encontradas: " + Chr(13) + Chr(13)
    
    
    Do While wResumo.Cells(x, 1) <> ""
    
       If wResumo.Cells(x, 2) <> "" Then
          cMensagem = cMensagem + "   [" + Str(wResumo.Cells(x, 3)) + " x ]" + wResumo.Cells(x, 2) + Chr(13)
       Else
          cMensagem = cMensagem + "   [" + Str(wResumo.Cells(x, 3)) + " x ]" + " Sem Função" + Chr(13)
       End If
      
       nSumOcor = nSumOcor + wResumo.Cells(x, 3)
       
       x = x + 1
       
       If x >= 28 And nPagina = 0 Then
          nPagina = 1
          cMensagem1 = cMensagem
          cMensagem = "Busca:" + wResumo.Cells(2, 11) + Chr(13) + Chr(13) + "Funções Encontradas: " + Chr(13) + Chr(13)
       End If
          
    Loop
    
    If x = 3 Then
       cMensagem = cMensagem + Chr(13) + " Nenhuma ocorrência "
    End If
    
    nSumFunc = x - 3
    
    If cMensagem1 <> "" Then
       MsgBox (cMensagem1 + Chr(13))
    End If
    
    MsgBox (cMensagem + Chr(13) + "Func:" + Str(nSumFunc) + " / Ocorr: " + Str(nSumOcor))
End If


wResumo.Cells(1, 1) = Now()
    
End Sub

