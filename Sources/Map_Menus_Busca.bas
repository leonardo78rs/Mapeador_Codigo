Sub BuscaMenus()

Dim cFile As String

Dim cBusca As String
Dim cPath As String
Dim na As Double
Dim nb As Double
Dim nc As Double
Dim nfontes As Double

Dim wFontes As Worksheet
Dim wMenus  As Worksheet
Dim cPathFile As String


Dim strTextLine
Dim strTextFile
Dim intFileNumber

Dim nLinhaFontes  As Integer
Dim nLoop         As Integer
Dim cNomeFonte    As String
Dim cTimeFim As String
Dim nTime    As Double
Dim lexibemsg As Boolean
Dim progresso As Double
Dim JaMostrouMensagem As Boolean
Dim nArquivosNaoEncontrados As Integer

Dim nArrayRel, nMenu, nConteudo As Integer
Dim cArrayRel, cMenu, cConteudo As String
Dim nLinRelatMenu, nLinRelatConteudo, nLinRelatRelat, nLinRelat As Integer

Dim tmp As String

Dim nArrayRelIni, nMenuIni, nConteudoIni As Integer
Dim nArrayRelFim, nMenuFim, nConteudoFim As Integer

Set wFontes = Sheets("Fontes")
Set wMenus = Sheets("Menus")
 
nLinhaFontes = 0
cPathFile = ""

intFileNumber = 1
nLinRelatMenu = 7
nLinRelatConteudo = 7
nLinRelatRelat = 7



For nLoop = 1 To 2000
' colComent = InStr(1, strTextLine, "/*", vbTextCompare)

    If InStr(1, wFontes.Cells(nLoop, 1), "menusiat") > 0 Then
     '   MsgBox ("encontrado")
       nLinhaFontes = nLoop
       cPathFile = wFontes.Cells(nLoop, 1)
       Exit For
    End If
Next nLoop

strTextFile = cPathFile

Open strTextFile For Input As #intFileNumber 'Criar conexão com o arquivo txt

wMenus.Cells(nLinRelatMenu, 2) = "Menu"
wMenus.Cells(nLinRelatMenu, 3) = "Descrição"
wMenus.Cells(nLinRelatMenu, 4) = "Função"


'Loop para percorrer as linhas do arquivo até o seu final
Do While Not EOF(intFileNumber)
   Line Input #intFileNumber, strTextLine
   LinFonte = LinFonte + 1
   
   ' AADD( aMenuDAJ, { "C", " C - Taxas para o Arquivo do Estoque                ", "", { || CCRPFJ() }, FECHA_ARQUIVOS } )
   ' Array de Menus no SIRETFU2.PRG
   ' AAdd( aConteudo, { 0085, "IMPORT. DADOS ENQUADR.CRON.E SALDOS BNDES/FCO/PROCAP"  , 001, "{|| RTADAO() }  "
 

   nMenu = 0
   nMenu = InStr(1, strTextLine, " AADD( aMenu", vbTextCompare)
   
   
   If nMenu > 0 Then
     
        nLinRelatMenu = nLinRelatMenu + 1
        nLinRelat = nLinRelatMenu
        nMenuIni = nMenu + 12    'prox aspa dupla
        
        
     
          
        If InStr(1, strTextLine, "X - Retorna", vbTextCompare) = 0 Then
        
           strTextLine = ConverteNaMao(strTextLine)
        
           If Mid(strTextLine, nMenuIni, 1) = "s" Then                                   'menus principais
              nMenuIni = InStr(nMenuIni, strTextLine, Chr(34), vbTextCompare) + 1
              nMenuFim = InStr(nMenuIni + 1, strTextLine, Chr(34), vbTextCompare) + 1
              wMenus.Cells(nLinRelat, 2) = (Mid(strTextLine, nMenuIni, nMenuFim - nMenuIni - 1))
           Else
           
              nMenuFim = InStr(nMenuIni + 1, strTextLine, ",", vbTextCompare) + 1
              wMenus.Cells(nLinRelat, 2) = Mid(strTextLine, nMenuIni, nMenuFim - nMenuIni - 1)
                           
              nMenuIni = InStr(nMenuFim, strTextLine, Chr(34), vbTextCompare) + 1        'prox aspa dupla
              nMenuFim = InStr(nMenuIni + 1, strTextLine, Chr(34), vbTextCompare) + 1
              wMenus.Cells(nLinRelat, 2) = wMenus.Cells(nLinRelat, 2) + "-" + (Mid(strTextLine, nMenuIni, nMenuFim - nMenuIni - 1))
              
              
              nMenuIni = InStr(nMenuFim, strTextLine, ",", vbTextCompare) + 8        'prox virgula e aspa dupla
              nMenuFim = InStr(nMenuIni + 1, strTextLine, ",", vbTextCompare) + 2
              wMenus.Cells(nLinRelat, 3) = Trim(Mid(strTextLine, nMenuIni, nMenuFim - nMenuIni - 3))
              
              tmp = Mid(wMenus.Cells(nLinRelat, 3), Len(wMenus.Cells(nLinRelat, 3)), 1)
              If tmp = Chr(16) Then
                 wMenus.Cells(nLinRelat, 3) = Trim(Mid(wMenus.Cells(nLinRelat, 3), 1, Len(wMenus.Cells(nLinRelat, 3)) - 1))
              End If
                      
              nMenuIni = InStr(nMenuFim, strTextLine, "{", vbTextCompare) + 5        'proxima {||
              nMenuFim = InStr(nMenuIni + 1, strTextLine, "}", vbTextCompare) + 1
              wMenus.Cells(nLinRelat, 4) = (Mid(strTextLine, nMenuIni, nMenuFim - nMenuIni - 1))
              
           End If
           
        Else
        
            nLinRelatMenu = nLinRelatMenu - 1
            
        End If
    
        
    End If
    
      
    Loop
   
'Fechar a conexão com o arquivo
Close #intFileNumber


nLinRelatMenu = 7
nLinRelatConteudo = 7
nLinRelatRelat = 7
nLinRelat = 7

For nLoop = 1 To 2000
' colComent = InStr(1, strTextLine, "/*", vbTextCompare)

    If InStr(1, wFontes.Cells(nLoop, 1), "siretfu2") > 0 Then     'siretfu2    menusiat
     '   MsgBox ("encontrado")
       nLinhaFontes = nLoop
       cPathFile = wFontes.Cells(nLoop, 1)
       Exit For
    End If
Next nLoop

strTextFile = cPathFile

Open strTextFile For Input As #intFileNumber 'Criar conexão com o arquivo txt

wMenus.Cells(nLinRelatMenu, 2) = "Menu"
wMenus.Cells(nLinRelatMenu, 3) = "Descrição"
wMenus.Cells(nLinRelatMenu, 4) = "Função"


'Loop para percorrer as linhas do arquivo até o seu final
Do While Not EOF(intFileNumber)
   Line Input #intFileNumber, strTextLine
   LinFonte = LinFonte + 1
   
   ' AADD( aMenuDAJ, { "C", " C - Taxas para o Arquivo do Estoque                ", "", { || CCRPFJ() }, FECHA_ARQUIVOS } )
   ' Array de Menus no SIRETFU2.PRG
   ' AAdd( aConteudo, { 0085, "IMPORT. DADOS ENQUADR.CRON.E SALDOS BNDES/FCO/PROCAP"  , 001, "{|| RTADAO() }  "
   
   nMenu = 0
   nMenu = InStr(1, strTextLine, " AADD( aMenu", vbTextCompare)
   
   
   If nMenu > 0 Then
     
        nLinRelatMenu = nLinRelatMenu + 1
        nLinRelat = nLinRelatMenu
        nMenuIni = nMenu + 12    'prox aspa dupla
        
        
        If InStr(1, strTextLine, "X - Retorna", vbTextCompare) = 0 Then
        
           strTextLine = ConverteNaMao(strTextLine)
           
           If Mid(strTextLine, nMenuIni, 1) = "s" Then                                   'menus principais
              nMenuIni = InStr(nMenuIni, strTextLine, Chr(34), vbTextCompare) + 1
              nMenuFim = InStr(nMenuIni + 1, strTextLine, Chr(34), vbTextCompare) + 1
              wMenus.Cells(nLinRelat, 7) = (Mid(strTextLine, nMenuIni, nMenuFim - nMenuIni - 1))
           Else
           
              nMenuFim = InStr(nMenuIni + 1, strTextLine, ",", vbTextCompare) + 1
              wMenus.Cells(nLinRelat, 7) = Mid(strTextLine, nMenuIni, nMenuFim - nMenuIni - 1)
                           
              nMenuIni = InStr(nMenuFim, strTextLine, Chr(34), vbTextCompare) + 1        'prox aspa dupla
              nMenuFim = InStr(nMenuIni + 1, strTextLine, Chr(34), vbTextCompare) + 1
              wMenus.Cells(nLinRelat, 7) = wMenus.Cells(nLinRelat, 7) + "-" + (Mid(strTextLine, nMenuIni, nMenuFim - nMenuIni - 1))
              
              
              nMenuIni = InStr(nMenuFim, strTextLine, ",", vbTextCompare) + 8        'prox virgula e aspa dupla
              nMenuFim = InStr(nMenuIni + 1, strTextLine, ",", vbTextCompare) + 2
              wMenus.Cells(nLinRelat, 8) = Trim(Mid(strTextLine, nMenuIni, nMenuFim - nMenuIni - 3))
              
              tmp = Mid(wMenus.Cells(nLinRelat, 8), Len(wMenus.Cells(nLinRelat, 8)), 1)
              If tmp = Chr(16) Then
                 wMenus.Cells(nLinRelat, 8) = Trim(Mid(wMenus.Cells(nLinRelat, 8), 1, Len(wMenus.Cells(nLinRelat, 8)) - 1))
              End If
                      
              nMenuIni = InStr(nMenuFim, strTextLine, "{", vbTextCompare) + 5        'proxima {||
              nMenuFim = InStr(nMenuIni + 1, strTextLine, "}", vbTextCompare) + 1
              wMenus.Cells(nLinRelat, 9) = (Mid(strTextLine, nMenuIni, nMenuFim - nMenuIni - 1))
              
           End If
           
        Else
        
            nLinRelatMenu = nLinRelatMenu - 1
            
        End If
    
        
    End If
    
    
    '  AADD( aConteudoDAJ, { "C", " C - Taxas para o Arquivo do Estoque                ", "", { || CCRPFJ() }, FECHA_ARQUIVOS } )
    ' AAdd( aConteudo, { 0085, "IMPORT. DADOS ENQUADR.CRON.E SALDOS BNDES/FCO/PROCAP"  , 001, "{|| RTADAO() }  "
    ' AAdd( aConteudo, { 001, 00, "FCUASLOTE_A" , "C", Space( 80 )    , "UAs Proc. Abertura/Final de Dia lote A"
 
    nConteudo = 0
    nConteudo = InStr(1, strTextLine, " AADD( aConteudo", vbTextCompare)
   
   
    If nConteudo > 0 And InStr(1, strTextLine, "AAdd( aConteudo, { 001, ", vbTextCompare) = 0 Then         'excluir uns parametros
     
       strTextLine = ConverteNaMao(strTextLine)
       
       nLinRelatConteudo = nLinRelatConteudo + 1
       nLinRelat = nLinRelatConteudo
        
       '  nConteudoIni = nConteudo + 12    'prox aspa dupla
          
        
       nConteudoIni = InStr(1, strTextLine, "{", vbTextCompare) + 1
       nConteudoFim = InStr(nConteudoIni + 1, strTextLine, ",", vbTextCompare) + 1
       wMenus.Cells(nLinRelat, 12) = Mid(strTextLine, nConteudoIni, nConteudoFim - nConteudoIni - 1)
                    
       nConteudoIni = InStr(nConteudoFim, strTextLine, Chr(34), vbTextCompare) + 1        'prox aspa dupla
       nConteudoFim = InStr(nConteudoIni + 1, strTextLine, Chr(34), vbTextCompare) + 1
       wMenus.Cells(nLinRelat, 13) = (Mid(strTextLine, nConteudoIni, nConteudoFim - nConteudoIni - 1))
       
       nConteudoIni = InStr(nConteudoFim, strTextLine, "{", vbTextCompare) + 4       'prox virgula e aspa dupla
       nConteudoFim = InStr(nConteudoIni + 1, strTextLine, "}", vbTextCompare) - 1
       wMenus.Cells(nLinRelat, 14) = Trim(Mid(strTextLine, nConteudoIni, nConteudoFim - nConteudoIni - 0))
       
       tmp = Mid(wMenus.Cells(nLinRelat, 14), Len(wMenus.Cells(nLinRelat, 14)), 1)
       
       If tmp = Chr(16) Then
          wMenus.Cells(nLinRelat, 14) = Trim(Mid(wMenus.Cells(nLinRelat, 14), 1, Len(wMenus.Cells(nLinRelat, 14)) - 1))
       End If
               

    End If
     
     
   ' AADD( aArrayRel, {"001", "0001", 3, "PC", "RELATORIO DE MOVIMENTOS NAO LANCADOS NO CONTA CORRENTE"} )
   ' AADD( aArrayRel, {"001", "0001", 4, "PD", "RELATORIO AMORTIZACOES POR BORDERO DE LIBERACAO" } )
   
   ' Array de Relatórios no SIRETFU2.PRG
 
   
    nArrayRel = 0
    nArrayRel = InStr(1, strTextLine, " AADD( aArrayRel", vbTextCompare)
   
    If nArrayRel > 0 Then
    
       strTextLine = ConverteNaMao(strTextLine)
       
       nLinRelatRelat = nLinRelatRelat + 1
       nLinRelat = nLinRelatRelat
       
       nArrayRelIni = InStr(nArrayRel, strTextLine, Chr(34), vbTextCompare) + 1        'prox aspa dupla
       nArrayRelFim = InStr(nArrayRelIni + 1, strTextLine, Chr(34), vbTextCompare) + 1
       wMenus.Cells(nLinRelat, 17) = Mid(strTextLine, nArrayRelIni, nArrayRelFim - nArrayRelIni - 1)
       
            
       nArrayRelIni = InStr(nArrayRelFim, strTextLine, Chr(34), vbTextCompare) + 1        'prox aspa dupla
       nArrayRelFim = InStr(nArrayRelIni + 1, strTextLine, Chr(34), vbTextCompare) + 1
       wMenus.Cells(nLinRelat, 18) = (Mid(strTextLine, nArrayRelIni, nArrayRelFim - nArrayRelIni - 1))
          
       nArrayRelIni = InStr(nArrayRelFim, strTextLine, ",", vbTextCompare) + 1        'prox aspa dupla
       nArrayRelFim = InStr(nArrayRelIni + 1, strTextLine, ",", vbTextCompare) + 1
       wMenus.Cells(nLinRelat, 19) = (Mid(strTextLine, nArrayRelIni, nArrayRelFim - nArrayRelIni - 1))

              
       nArrayRelIni = InStr(nArrayRelFim, strTextLine, Chr(34), vbTextCompare) + 1        'prox aspa dupla
       nArrayRelFim = InStr(nArrayRelIni + 1, strTextLine, Chr(34), vbTextCompare) + 1
       wMenus.Cells(nLinRelat, 20) = (Mid(strTextLine, nArrayRelIni, nArrayRelFim - nArrayRelIni - 1))

           
       nArrayRelIni = InStr(nArrayRelFim, strTextLine, Chr(34), vbTextCompare) + 1        'prox aspa dupla
       nArrayRelFim = InStr(nArrayRelIni + 1, strTextLine, Chr(34), vbTextCompare) + 1
       wMenus.Cells(nLinRelat, 21) = (Mid(strTextLine, nArrayRelIni, nArrayRelFim - nArrayRelIni - 1))
 
    End If
      
    Loop
   
'Fechar a conexão com o arquivo
Close #intFileNumber

ClassifMenuSiac


End Sub


