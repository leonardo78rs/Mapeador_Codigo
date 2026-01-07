Sub Estatica()
Dim wFontes As Worksheet
Dim wOcorr As Worksheet
Dim numfontes  As Integer
Dim wResumo As Worksheet
Dim n As Integer

Set wFontes = Sheets("Fontes")
Set wOcorr = Sheets("Ocorrencias")
Set wResumo = Sheets("Resumo")

n = 10
'wOcorr.Cells(2, 2) = "Function"  ' -- Palavra a buscar
'Principal

'wOcorr.Cells(n, 26) = wOcorr.Cells(3, 2) '-- PASTA
'wOcorr.Cells(n, 27) = wOcorr.Cells(7, 2) '-- qtde ARQUIVOS fontes
'wOcorr.Cells(n, 28) = wOcorr.Cells(10, 2) '-- qtde functions // qtde RESULTADOS
'wOcorr.Cells(n, 31) = wOcorr.Cells(4, 2) '-- qtde linhas


'wOcorr.Cells(2, 2) = "Procedure"  ' -- Palavra a buscar
'Principal
'wOcorr.Cells(n, 29) = wOcorr.Cells(10, 2) '-- qtde procedure / qtde RESULTADOS


'wOcorr.Cells(2, 2) = "Infra_HTTPClient"
'Principal
'wOcorr.Cells(n, 32) = wOcorr.Cells(10, 2) '-- qtde chamadas API // qtde RESULTADOS


'wOcorr.Cells(2, 2) = "F_Facade"
'Principal
'wOcorr.Cells(n, 33) = wOcorr.Cells(10, 2) '-- qtde chamadas plsql // qtde RESULTADOS


wOcorr.Cells(2, 2) = "Infra_HTTPSOAP"
Principal
wOcorr.Cells(n, 34) = wOcorr.Cells(10, 2) '-- qtde chamadas SOAP // qtde RESULTADOS
wOcorr.Cells(2, 2) = "AcessoSoap"
Principal
wOcorr.Cells(n, 34) = wOcorr.Cells(n, 34) + wOcorr.Cells(10, 2) '-- qtde chamadas SOAP // qtde RESULTADOS




End Sub


