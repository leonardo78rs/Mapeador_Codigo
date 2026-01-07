Sub LimpaOcorrencias()

Dim nLimpa As Integer
Dim wOcorr As Worksheet
Dim nEmBranco As Integer

Set wOcorr = Sheets("Ocorrencias")
nEmBranco = 0

For nLimpa = 13 To 10000
    wOcorr.Cells(nLimpa, 1) = ""
    wOcorr.Cells(nLimpa, 2) = ""
    wOcorr.Cells(nLimpa, 3) = ""
    wOcorr.Cells(nLimpa, 4) = ""
    wOcorr.Cells(nLimpa, 5) = ""
    
    If wOcorr.Cells(nLimpa + 1, 1) = "" Then
       nEmBranco = nEmBranco + 1
    End If
    
    If nEmBranco > 3 Then
       nLimpa = 10000         'Força a sair do loop
    End If
    
Next
    

End Sub


