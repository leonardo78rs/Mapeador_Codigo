Sub DesenhaFormas()
' Pré-requisito: Execucao da Sub Arvore


Dim wArvore As Worksheet
Dim wDesenho As Worksheet
Dim nOid As Integer
Dim nOidPai As Integer
Dim Const_Larg_Caixa, Const_Altu_Caixa As Integer

 
Set wArvore = Sheets("Arvore")
Set wDesenho = Sheets("Desenho")
 
 
ClassificaObjetos
Sheets("Desenho").Select
Cells.Select
Selection.Delete Shift:=xlUp

wDesenho.Activate
Columns("A:ZZ").Select
Selection.ClearContents

'Definicoes constantes
Const_Larg_Caixa = 200
Const_Altu_Caixa = 20
Const_Meio_Caixa = 10
Const_TotalProcessar = 2000
Const_Posicao_Linha_Normal = 0.75

'Colunas guia Arvore R..Z
Def_ID = 18
Def_IDPai = 19
Def_Funcao = 20
Def_Tipo = 21
Def_CaixaLin = 22
Def_CaixaCol = 23
Def_CaixaPaiLin = 24
Def_CaixaPaiCol = 25
Def_DeslocLin = 26


For nLoop = 2 To Const_TotalProcessar

    EhCorrespondenciaMultipla = False
    DeslocaSetaVerticalmente = 0

    If wArvore.Cells(nLoop, Def_ID) = Empty Then
       Exit For
    End If
    
    'Se não exise ainda, cria a caixa
    If (wArvore.Cells(nLoop, Def_ID) <> wArvore.Cells(nLoop - 1, Def_ID)) Then
       
       a = DesenhaCaixas(wArvore.Cells(nLoop, Def_CaixaLin), _
                         wArvore.Cells(nLoop, Def_CaixaCol), _
                         200, _
                         20, _
                         wArvore.Cells(nLoop, Def_Funcao), _
                         wArvore.Cells(nLoop, Def_Tipo))
                         
    End If
    
    'Se for caixa-filha, coloca o conector para a caixa-pai
    If wArvore.Cells(nLoop, Def_IDPai) <> Empty Then
    
        'Trocar a cor de acordo com a caixa-alvo
        Dif_Cor_e_DeslocHorizontal = 0.7 - (wArvore.Cells(nLoop, Def_CaixaPaiLin) / (20 * wArvore.Cells(1, 29)))
           
        If Dif_Cor_e_DeslocHorizontal < 0.2 Then
           aCores = Array(255, 0, 0)    'vermelho
        ElseIf Dif_Cor_e_DeslocHorizontal < 0.4 Then
           aCores = Array(240, 200, 20)  'amarelo
        ElseIf Dif_Cor_e_DeslocHorizontal < 0.5 Then
           aCores = Array(0, 255, 0)    'verde
        ElseIf Dif_Cor_e_DeslocHorizontal < 0.6 Then
            aCores = Array(200, 50, 200) 'violeta
        Else
            aCores = Array(255, 50, 0)  'laranja
        End If
    
    
        'Posicao linha normal 0.75  (1 pra 1 ou N pra 1)
        'Posicao linha menor --> 1 pra N
        If wArvore.Cells(nLoop, Def_DeslocLin) <> Const_Posicao_Linha_Normal Then
        
           EhCorrespondenciaMultipla = True
           DeslocaSetaVerticalmente = 1
           'Casos onde coincide várias setas chegando em um alvo -- cada cor chega a uma altura da caixa
           DeslocaSetaVerticalmente = DeslocaSetaVerticalmente + (5 * Dif_Cor_e_DeslocHorizontal)
           
        End If


        ActiveSheet.Shapes.AddConnector(msoConnectorElbow, _
                                        wArvore.Cells(nLoop, Def_CaixaCol), _
                                        wArvore.Cells(nLoop, Def_CaixaLin) + Const_Meio_Caixa + DeslocaSetaVerticalmente, _
                                        wArvore.Cells(nLoop, Def_CaixaPaiCol) + Const_Larg_Caixa, _
                                        wArvore.Cells(nLoop, Def_CaixaPaiLin) + Const_Meio_Caixa + DeslocaSetaVerticalmente).Select
                                        
        Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadTriangle  'seta do conector
        
       
        If EhCorrespondenciaMultipla = True Then
           'Casos onde coincide várias setas chegando em um alvo
           'troca cor
           Selection.ShapeRange.Line.ForeColor.RGB = RGB(aCores(0), aCores(1), aCores(2))
           'desloca horiz. a quebra do conector
           Selection.ShapeRange.Adjustments.Item(1) = Dif_Cor_e_DeslocHorizontal
           
        Else
        
           'Casos Normais: Seta Azul e sempre quebrando em 75% do tamanho do conector
           Selection.ShapeRange.Adjustments.Item(1) = wArvore.Cells(nLoop, Def_DeslocLin)
              
        End If
        
       
      
        
    End If
    
Next nLoop



'funcao principal
 
 
 a = DesenhaCaixas(40, 40, 200, Const_Altu_Caixa * 3, wArvore.Cells(1, 1), "prg")
 
 
 linhaarvore = 2
 
 Do While (wArvore.Cells(linhaarvore, Def_Funcao) <> "")
    
    If wArvore.Cells(linhaarvore, Def_IDPai) = Empty Then
    
        ActiveSheet.Shapes.AddConnector(msoConnectorElbow, _
                                        wArvore.Cells(linhaarvore, Def_CaixaCol), _
                                        wArvore.Cells(linhaarvore, Def_CaixaLin) + Const_Meio_Caixa, _
                                        40 + Const_Larg_Caixa, _
                                        40 + Const_Meio_Caixa).Select
                                        
        Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadTriangle
       
    End If
    
    linhaarvore = linhaarvore + 1
    
 Loop

End Sub
