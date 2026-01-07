Function ImportTxtFile(cPath As String, cFile As String, cBusca As String, LinOcorrencias As Double) As Double

Dim strTextLine
Dim strTextFile
Dim intFileNumber
Dim wOcorr As Worksheet
Dim wPlan2 As Worksheet
Dim lin, col, colFunc As Integer
Dim LinFonte As Double
Dim ncoment As Integer
Dim ncomentSql As Integer

Dim cFuncAtual As String
Dim nini, nfim As Integer
Dim colComent As Integer
Dim lBlocoComentado As Boolean


Set wOcorr = Sheets("Ocorrencias")

lBlocoComentado = False
lin = LinOcorrencias
LinFonte = 0
intFileNumber = 1  'Criar numeração
strTextFile = cPath + cFile
cFuncAtual = ""

lIgnoraComentario = wOcorr.Cells(8, 4)



Open strTextFile For Input As #intFileNumber 'Criar conexão com o arquivo txt

'Loop para percorrer as linhas do arquivo até o seu final
Do While Not EOF(intFileNumber)
   Line Input #intFileNumber, strTextLine
   LinFonte = LinFonte + 1
   
   ' wOcorr.Cells(2, 30) = wOcorr.Cells(2, 30) + 1   'TEMPORARIO -- SO PRA CONTAR QUANTAS LINHAS (EM TORNO DE 1.2 MILHOES LINHAS  CLIPPER E PL)
   
   colComent = 0
   colComent = InStr(1, strTextLine, "/*", vbTextCompare)
   colComent2 = InStr(1, strTextLine, "*/", vbTextCompare)
   
   If colComent > 0 And colComent2 = 0 Then lBlocoComentado = True
   If lBlocoComentado And colComent = 0 And colComent2 > 0 Then lBlocoComentado = False
   
   '-- Se for comentada a linha, ou parte do bloco comentado, não lê a linha
   If Not lIgnoraComentario Or Not lBlocoComentado Then
   
        If (colComent = 0) Or (colComent > 0 And colComent2 > 0) Then
        
             If Mid(Trim(strTextLine), 1, 1) <> "/" Or (Not lIgnoraComentario) Then
             
                ncoment = 0
                If lIgnoraComentario Then
                   ncoment = InStr(1, strTextLine, "//", vbTextCompare)
                
                     
                     If ncoment > 0 Then
                        strTextLine = Mid(strTextLine, 1, ncoment - 1)         ' se encontrou //, corta tudo que vem depois
                     Else
                        ncomentSql = 0
                        ncomentSql = InStr(1, strTextLine, "--", vbTextCompare)
                        If ncoment > 0 Then
                           strTextLine = Mid(strTextLine, 1, ncomentSql - 1)   ' se encontrou --, corta tudo que vem depois
                        End If
                     End If
                
                End If
                
                colFunc = 0
                colProc = 0
                colComent = 0
                
                colFunc = InStr(1, strTextLine, "function", vbTextCompare)
                
                If lBlocoComentado Then
                   colComent = InStr(1, strTextLine, "Funcao ........: ", vbTextCompare)
                End If
                
                colProc = InStr(1, strTextLine, "Procedure", vbTextCompare)
                
                If colFunc = 0 And colProc <> 0 Then
                   colFunc = colProc
                End If
                
                If colFunc = 0 And colComent <> 0 Then
                   colFunc = colComent
                End If
                
                If colFunc <> 0 Then
                   nini = 0
                   nfim = 0
                   nini = InStr(1, strTextLine, "function", vbTextCompare) + 8
                   If colProc <> 0 Then
                      nini = InStr(1, strTextLine, "Procedure", vbTextCompare) + 9
                   End If
                                      
                   nfim = InStr(nini, strTextLine, "(", vbTextCompare)
                  
                   If nini > 0 And (nfim - nini) >= 0 Then
                      cFuncAtual = Mid(strTextLine, nini, (nfim - nini) + 1) + ")"
                   ElseIf colComent <> 0 Then
                      nini = InStr(1, strTextLine, "Funcao ........: ", vbTextCompare) + 17
                      ' nfim = Len(strTextLine)    'Len(Mid(strTextLine, nini, Len(strTextLine)))
                      cFuncAtual = " " + Mid(strTextLine, nini, Len(strTextLine)) + "()"
                   End If

                   
                End If
                    
                col = 0
                col = InStr(1, strTextLine, cBusca, vbTextCompare)
                
                
                cBuscaSemParentesis = Replace(cBusca, "(", "")
                cBuscaSemParentesis = Replace(cBuscaSemParentesis, ")", "")
                
                'nao trazer a propria funcao
                'nao trazer tb nos casos de Pl/sql que tem entre aspas e o End function
                
                nao_eh_a_propria_funcao = InStr(1, UCase(strTextLine), UCase("unction " + Trim(cBusca)), vbTextCompare) + _
                                          InStr(1, UCase(strTextLine), UCase("'" + Trim(cBuscaSemParentesis) + "'"), vbTextCompare) + _
                                          InStr(1, UCase(strTextLine), UCase("end " + Trim(cBuscaSemParentesis)), vbTextCompare)
                                         
                
                
                '--- TRATAMENTO MENUS
                If Trim(cFuncAtual) = "F_PMnuProSiat()" Or Trim(cFuncAtual) = "F_PMnuMovSiat()" Or Trim(cFuncAtual) = "F_PMnuRelSiat()" Then
                   cFuncAtual = "MENU SIAC"
                End If
 
                If Trim(cFuncAtual) = "FP_ContRot()" Or Trim(cFuncAtual) = "FP_MenuRotSiret()" Or Trim(cFuncAtual) = "FP_MenuRelSiret()" Then
                   cFuncAtual = "ROTINA (SIRET)"
                End If
                                                    
                                          
                                          
                If col <> 0 And nao_eh_a_propria_funcao = 0 Then
                   lin = lin + 1
                   wOcorr.Cells(lin, 1) = cFile
                   wOcorr.Cells(lin, 2) = LinFonte
                   wOcorr.Cells(lin, 3) = col
                   If cFuncAtual = "" Then
                      wOcorr.Cells(lin, 4) = " Função Geral "
                   Else
                      wOcorr.Cells(lin, 4) = cFuncAtual
                   End If
                   
         

                   
                   
                   wOcorr.Cells(lin, 5) = strTextLine
                End If
             
             End If
        End If
     End If
     
Loop
ImportTxtFile = lin


wOcorr.Cells(4, 2) = wOcorr.Cells(4, 2) + LinFonte

'Fechar a conexão com o arquivo
Close #intFileNumber

End Function

