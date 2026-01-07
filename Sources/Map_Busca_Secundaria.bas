Sub Secundario()

Dim cFile As String
Dim cBusca As String
Dim cPath As String
Dim na As Double
Dim nb As Double
Dim nc As Double
Dim wFontes As Worksheet
Dim nfontes As Integer
Dim nLimpa  As Integer
Dim wResumo As Worksheet
Dim nBuscas As Integer
Dim nMaxBuscas As Integer
Dim wOcorr As Worksheet
Dim cTimeIni As String
Dim cTimeFim As String
Dim nTime    As Double
Dim nEmBranco As Integer
Dim nFonteEmBranco As Integer
Dim nResumEmBranco As Integer
Dim cchamador As String
Dim lexibemsg As Boolean

lexibemsg = False

Set wFontes = Sheets("Fontes")
Set wResumo = Sheets("Resumo")
Set wOcorr = Sheets("Ocorrencias")



cTimeIni = Time()
wOcorr.Cells(10, 15) = 0

nEmBranco = 0                   'Quando tiver n>3 em branco, sai do loop
nResumoEmBranco = 0

'cPath = wOcorr.Cells(3, 2)     'Ocorrencias!B3   -- Pasta que contém os arquivos
cPath = ""

'Copia de A:C para  K:M  apenas se tiverem resultados
' Bug de quando executa duas vezes a pesquisa secundária

If wResumo.Cells(3, 1) <> Empty Then

    For nLimpa = 2 To 200
        wResumo.Cells(nLimpa + 3, 11) = wResumo.Cells(nLimpa, 1)
        wResumo.Cells(nLimpa + 3, 12) = wResumo.Cells(nLimpa, 2)
        wResumo.Cells(nLimpa + 3, 13) = wResumo.Cells(nLimpa, 3)
    Next
    
End If

'  Limpar tudo da celula N para a direita
wResumo.Activate
Columns("N:N").Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.ClearContents
Columns("A:C").Select
Sheets("Ocorrencias").Select


For nBuscas = 6 To 90

    LimpaOcorrencias
    
    wResumo.Activate
    LimpaResumo ("ABC")

    cBusca = wResumo.Cells(nBuscas, 12)  'Coluna L
    cBusca = Trim(cBusca)
    
    If cBusca <> Empty Then   ' PERGUNTAR SE NAO É A PRIMARIA
    
       If Mid(cBusca, 1, Len(cBusca) - 2) <> wOcorr.Cells(2, 2) And _
          Trim(Mid(cBusca, 1, LenB(cBusca) - 2)) <> Trim(wOcorr.Cells(2, 2)) Then               ' Nao buscar a própria palavra Ocorrencias!B2
       
            cBusca = Mid(cBusca, 1, Len(cBusca) - 1)
            
            na = 12
            nFonteEmBranco = 0              'Quando tiver n>3 em branco, sai do loop
            
            For nfontes = 2 To 1200
            
                cFile = wFontes.Cells(nfontes, 1)
                If Len(Trim(cFile)) > 0 Then
                       
                       na = ImportTxtFile(cPath, cFile, cBusca, na)
                
                Else
                    
                    nFonteEmBranco = nFonteEmBranco + 1
                    
                    If nFonteEmBranco > 3 Then
                       nfontes = 1200         'Força a sair do loop
                    End If
                
                End If
            Next
            
            GroupFunc (False)
            
           
            'Coloca os resultados horizontalmente a partir da coluna Ó
            
            nColIni = 15 'coluna Ó
            For nAchados = 3 To 43   'ver ultima
            
                If Trim(wResumo.Cells(nAchados, 1)) = Empty Then
                   Exit For
                End If
            
                If wResumo.Cells(nAchados, 2) <> wResumo.Cells(nBuscas, 12) And _
                   wResumo.Cells(nAchados, 2) <> Empty Then        'nao colocar a propria funcao
        
                   wResumo.Cells(nBuscas, nColIni) = wResumo.Cells(nAchados, 2)
                   nColIni = nColIni + 1
                
                End If
            
            
            Next
            
        End If
     Else
        nEmBranco = nEmBranco + 1
        If nEmBranco > 3 Then
           nBuscas = 90        'Força a sair do loop
        End If
     End If

Next


cTimeFim = Time()

nTime = Round((Mid(cTimeFim, 1, 2) * 3600 + Mid(cTimeFim, 4, 2) * 60 + Mid(cTimeFim, 7, 2) - _
               Mid(cTimeIni, 1, 2) * 3600 - Mid(cTimeIni, 4, 2) * 60 - Mid(cTimeIni, 7, 2)), 2)
        
 
wOcorr.Cells(10, 15) = nTime

Arvore

End Sub

