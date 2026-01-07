Function ExtraiNomesFunc()

Dim wResumo As Worksheet
Dim lin1, lin2, col As Integer
Dim nini, nfim As Double

Set wResumo = Sheets("Resumo")

For lin1 = 1 To 600    'criar variavel ou celula para total
    nini = 0
    nfim = 0
    nini = InStr(1, wResumo.Cells(lin1, 5), "function", vbTextCompare) + 8
    If nini = 0 Then
       nini = InStr(1, wResumo.Cells(lin1, 5), "procedure", vbTextCompare) + 9
    End If
    
    
    nfim = InStr(nini, wResumo.Cells(lin1, 5), "(", vbTextCompare)
   
    If (nfim - nini) >= 0 Then
       wResumo.Cells(lin1, 4) = Mid(wResumo.Cells(lin1, 5), nini, (nfim - nini) + 1) + ")"
    End If
         
Next
ExtraiNomesFunc = True

End Function


