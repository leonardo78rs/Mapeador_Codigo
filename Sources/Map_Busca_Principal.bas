Sub Principal()
Dim cFile As String
Dim cBusca As String
Dim cPath As String
Dim na As Double
Dim nb As Double
Dim nc As Double
Dim nfontes As Double
Dim wFontes As Worksheet
Dim wOcorr As Worksheet
Dim numfontes  As Integer
Dim wResumo As Worksheet
Dim lAnimado As Boolean
Dim cTimeIni As String
Dim cTimeFim As String
Dim nTime    As Double
Dim lexibemsg As Boolean
Dim progresso As Double
Dim JaMostrouMensagem As Boolean
Dim nArquivosNaoEncontrados As Integer


nArquivosNaoEncontrados = 0
lexibemsg = False

Set wFontes = Sheets("Fontes")
Set wOcorr = Sheets("Ocorrencias")
Set wResumo = Sheets("Resumo")

cTimeIni = Time()
wOcorr.Cells(8, 9) = 0
wOcorr.Cells(4, 2) = 0 ' qtde linhas nos fontes

JaMostrouMensagem = False 'mensagem de erro ao ler arquivo

'cPath = wOcorr.Cells(3, 2)     'Ocorrencias!B3   -- Pasta que contém os arquivos
cPath = ""
cBusca = wOcorr.Cells(2, 2)    'Ocorrencias!B2   -- Palavra a buscar
lAnimado = wOcorr.Cells(6, 4)  'Indica se a tela ficará mostrando o que está acontecendo (demora mais)

numfontes = 0                  'quantidade de fontes pesquisados

If Not lAnimado Then
   wOcorr.Cells(6, 2) = ""
   wOcorr.Cells(7, 2) = numfontes
End If
   
na = 12   'retorno da outra funcao e linha inicial da planilha ocorrencias

LimpaOcorrencias

nFonteEmBranco = 0

For nfontes = 2 To 1200
       
       cFile = wFontes.Cells(nfontes, 1)
       numfontes = numfontes + 1
        
       'exibir no formulario o percentual
       'nao fica mostrando (trava e mostra ultimo valor)
       'If wFontes.Cells(6, 8) <> 0 Then
       '   progresso = (numfontes / wFontes.Cells(6, 8))
       '   progresso = 100 * Round(progresso, 0)
       '   UserForm1.Label11.Caption = "Progresso: " + Str(progresso) + "%"
       'End If
       
       If Len(Trim(cFile)) > 0 Then
       
          If lAnimado Then
             wOcorr.Cells(6, 2) = cFile ' Colocar em tela o fonte que está buscando...
             wOcorr.Cells(7, 2) = numfontes
          End If
          
          If Dir(cPath + cFile) <> vbNullString Then
             na = ImportTxtFile(cPath, cFile, cBusca, na)
          ElseIf Not JaMostrouMensagem Then
             MsgBox ("Arquivo não encontrado, verifique!" + Chr(13) + Chr(13) + cFile)
             JaMostrouMensagem = True
             nArquivosNaoEncontrados = nArquivosNaoEncontrados + 1
          Else
              nArquivosNaoEncontrados = nArquivosNaoEncontrados + 1
          End If
        
        Else
               
               nFonteEmBranco = nFonteEmBranco + 1
               
               If nFonteEmBranco > 3 Then
                  nfontes = 1200         'Força a sair do loop
               End If
           
        
       End If
 
Next

If Not lAnimado Then
  wOcorr.Cells(6, 2) = cFile ' Colocar em tela o fonte que está buscando...
  wOcorr.Cells(7, 2) = numfontes
End If

LimpaResumo ("ABC")

' Faz o resumo das funções
'GroupFunc(lAnimado, lexibemsg)
GroupFunc (lexibemsg)

LimpaResumo ("KLM")

cTimeFim = Time()

nTime = Round((Mid(cTimeFim, 1, 2) * 3600 + Mid(cTimeFim, 4, 2) * 60 + Mid(cTimeFim, 7, 2) - _
               Mid(cTimeIni, 1, 2) * 3600 - Mid(cTimeIni, 4, 2) * 60 - Mid(cTimeIni, 7, 2)), 2)
        
        
wOcorr.Cells(8, 9) = nTime

If nArquivosNaoEncontrados > 0 Then
   MsgBox ("Total de " + Str(nArquivosNaoEncontrados) + " arquivos não encontrados ")
End If

End Sub


