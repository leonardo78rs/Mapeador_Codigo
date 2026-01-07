'--------------------------------------------------------------------------
' Função para limpar resumo
' Colunas ABC e KLM precisam ser limpas na funcao principal e na secundária
' Mas em momentos diferentes.
' Na secundária, vai limpar uma vez a KLM e vai limpar várias vezes (loop) as colunas ABC.

Sub LimpaResumo(cColunasLimpar As String)

Dim nBuscas As Integer
Dim wResumo As Worksheet
Dim nLimpa  As Integer
Dim EmBranco As Integer

Set wResumo = Sheets("Resumo")

increm = 0 ' Limpar colunas A,B,C ou colunas K,L,M
nEmBranco = 0

If cColunasLimpar = "KLM" Then
   increm = 10
End If
 
For nBuscas = 3 To 200
    wResumo.Cells(nBuscas, increm + 1) = ""
    wResumo.Cells(nBuscas, increm + 2) = ""
    wResumo.Cells(nBuscas, increm + 3) = ""
    
    If wResumo.Cells(nBuscas + 1, increm + 1) = Empty Then
       nEmBranco = nEmBranco + 1
    End If
        
    If nEmBranco > 3 Then
       nBuscas = 500         'Força a sair do loop
    End If
    
Next
    

End Sub

