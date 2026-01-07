Sub Arvore()

Dim wArvore As Worksheet
Dim wResumo As Worksheet
Dim nId, nIdPai, Linha As Integer

 
Set wArvore = Sheets("Arvore")
Set wResumo = Sheets("Resumo")
 
wArvore.Select
Range("a1:AZ400").Select
'Selection.ClearContents

'string que foi procurada
wArvore.Cells(1, 1) = wResumo.Cells(2, 11)
LoopArvore = 0
LoopResumo = 5              'A lista dos fontes/dependencias inicia em Resumo[ K6 ]
coluna_a_mais = 0
Linha = 2
nId = 1
nIdPai = 0
PosQuadroPrincLin = 40
PosQuadroPrincCol = 340

PosQuadroSecunLin = 40
PosQuadroSecunCol = 700

'MultiplicadorColunaQuandoRepetir = 1
'wArvore.Cells(1, 50) = 0
'wArvore.Cells(1, 51) = ""

'limpar celular R ate AA
    Sheets("Arvore").Select
    ActiveWindow.SmallScroll Down:=-18
    Range("A2:AE300").Select
    Selection.ClearContents
        
    'Range("R2:AA3").Select
    'Range(Selection, Selection.End(xlDown)).Select
    'Range(Selection, Selection.End(xlDown)).Select
    'Selection.ClearContents
    Range("R1").Select


Do While (wResumo.Cells(LoopResumo, 12) <> "")

   LoopResumo = LoopResumo + 1
   LoopArvore = LoopArvore + 1
   coluna_a_mais = 1
   wArvore.Cells(1 + LoopArvore, 2) = wResumo.Cells(LoopResumo, 12)
   
   'procura as funções que chamam as funções que contém a busca, na guia resumo
   'porém quando há apenas a função que contém a busca, ele estava ignorando
   Do While (wResumo.Cells(LoopResumo, 14 + coluna_a_mais) <> "") Or _
            (coluna_a_mais = 1 And wResumo.Cells(LoopResumo, 12) <> "")
     
      'Na guia resumo, as funcoes-filha estao na horizontal
      'Na guia arvore estao na vertical
      wArvore.Cells(1 + coluna_a_mais + LoopArvore, 3) = wResumo.Cells(LoopResumo, 14 + coluna_a_mais)
      wArvore.Cells(1 + coluna_a_mais + LoopArvore, 12) = wResumo.Cells(LoopResumo, 12)
      wArvore.Cells(1 + coluna_a_mais + LoopArvore, 13) = wResumo.Cells(LoopResumo, 14 + coluna_a_mais)
      wArvore.Cells(1 + coluna_a_mais + LoopArvore, 14) = Right(wResumo.Cells(LoopResumo, 11), 3)
       
       
      '------------------------------------ CONSTROI DADOS P/CAIXAS DA COLUNA DO MEIO
      'verifica se já existe a primária
      Lincolunat = 2
      PrincExistente = False
      
      Do While (wArvore.Cells(Lincolunat + 1, 20) <> "")
         
         If wArvore.Cells(Lincolunat, 20) = wResumo.Cells(LoopResumo, 12) And _
            wArvore.Cells(Lincolunat, 19) = Empty Then
            
            PrincExistente = True
            
            wArvore.Cells(Linha, 18) = wArvore.Cells(Lincolunat, 18)
            wArvore.Cells(Linha, 19) = Empty
            wArvore.Cells(Linha, 20) = wArvore.Cells(Lincolunat, 20)
            wArvore.Cells(Linha, 21) = wArvore.Cells(Lincolunat, 21)
            wArvore.Cells(Linha, 22) = wArvore.Cells(Lincolunat, 22)
            wArvore.Cells(Linha, 23) = wArvore.Cells(Lincolunat, 23)
            wArvore.Cells(Linha, 24) = Empty
            wArvore.Cells(Linha, 25) = Empty
            wArvore.Cells(Linha, 26) = Empty
            nIdPai = wArvore.Cells(Lincolunat, 18)
            
            Exit Do
            
         End If
      
         Lincolunat = Lincolunat + 1
         
      Loop
      'fim verifica se já existe
      
      If PrincExistente = False Then
         wArvore.Cells(Linha, 18) = nId
         wArvore.Cells(Linha, 19) = Empty
         wArvore.Cells(Linha, 20) = wResumo.Cells(LoopResumo, 12)
         wArvore.Cells(Linha, 21) = Right(wResumo.Cells(LoopResumo, 11), 3)
         wArvore.Cells(Linha, 22) = PosQuadroPrincLin
         wArvore.Cells(Linha, 23) = PosQuadroPrincCol
         wArvore.Cells(Linha, 24) = Empty
         wArvore.Cells(Linha, 25) = Empty
         wArvore.Cells(Linha, 26) = Empty
     
         nIdPai = nId
         nId = nId + 1
      End If
      
      
      Linha = Linha + 1
      
      
      '------------------------------------ CONSTROI DADOS P/CAIXAS DA COLUNA DA DIREITA
      
      'caso especial: se a primária está sozinha (sem filhas)
      If (coluna_a_mais = 1 And wResumo.Cells(LoopResumo, 12) <> "" And wResumo.Cells(LoopResumo, 12) <> "" And wResumo.Cells(LoopResumo, 14 + coluna_a_mais) = "") Then
         
         nId = nId + 1
         PosQuadroSecunLin = PosQuadroSecunLin + 30
         PosQuadroPrincLin = PosQuadroPrincLin + 30
         Exit Do
      
      End If
      
      
      'verifica se já existe secundario
      Lincolunat = 2
      SecundExistente = False
      
      Do While (wArvore.Cells(Lincolunat + 1, 20) <> "")
         
         If wArvore.Cells(Lincolunat, 20) = wResumo.Cells(LoopResumo, 14 + coluna_a_mais) And _
            wArvore.Cells(Lincolunat, 19) <> Empty Then
            
            SecundExistente = True
            
            wArvore.Cells(Linha, 18) = nId
            wArvore.Cells(Linha, 19) = wArvore.Cells(Lincolunat, 19)
            wArvore.Cells(Linha, 20) = wArvore.Cells(Lincolunat, 20)
            wArvore.Cells(Linha, 21) = wArvore.Cells(Lincolunat, 21)
            wArvore.Cells(Linha, 22) = wArvore.Cells(Lincolunat, 22) 'PosQuadroSecunLin
            wArvore.Cells(Linha, 23) = wArvore.Cells(Lincolunat, 23) 'PosQuadroSecunCol
            wArvore.Cells(Linha, 24) = wArvore.Cells(Linha - 1, 22)
            wArvore.Cells(Linha, 25) = wArvore.Cells(Linha - 1, 23)
            wArvore.Cells(Linha, 26) = 0.7 '''' "=int((0.7 - (" + Str(Linha) + " / (1.4 * $ac$1))))" ' wArvore.Cells(1, 29)))
            nIdPai = wArvore.Cells(Lincolunat, 18)
            
            Exit Do
            
         End If
      
         Lincolunat = Lincolunat + 1
         
      Loop
      'fim verifica se já existe secundario
      
      If SecundExistente = False Then
         wArvore.Cells(Linha, 18) = nId
         wArvore.Cells(Linha, 19) = nIdPai
         wArvore.Cells(Linha, 20) = wResumo.Cells(LoopResumo, 14 + coluna_a_mais)
         wArvore.Cells(Linha, 21) = Right(wResumo.Cells(LoopResumo, 11), 3)
         wArvore.Cells(Linha, 22) = PosQuadroSecunLin
         wArvore.Cells(Linha, 23) = PosQuadroSecunCol
         wArvore.Cells(Linha, 24) = wArvore.Cells(Linha - 1, 22)
         wArvore.Cells(Linha, 25) = wArvore.Cells(Linha - 1, 23)
         wArvore.Cells(Linha, 26) = 0.75
      
      End If
      
      Linha = Linha + 1
          
      '------------------------------------ FIM CAIXAS DA COLUNA DA DIREITA
      
      nId = nId + 1
      PosQuadroSecunLin = PosQuadroSecunLin + 30
      PosQuadroPrincLin = PosQuadroPrincLin + 30
      
      coluna_a_mais = coluna_a_mais + 1
      
      
   Loop
   
   LoopArvore = LoopArvore + coluna_a_mais - 1
   
Loop

'wArvore.Cells(1, 5) = Now()
' ClassifArvore

MsgBox ("desenhaformas desabilitado tmp")

'DesenhaFormas

End Sub
