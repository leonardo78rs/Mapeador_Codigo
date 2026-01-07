Sub Instrucoes()

Dim Ocorr As Worksheet
Set wOcorr = Sheets("Ocorrencias")


'MsgBox ("- Tem que copiar todos arquivos para a pasta indicada na celula B3: (" + wOcorr.Cells(3, 2) + ")" + Chr(10) + Chr(10) + _
'  "- Tem que colocar a lista de arquivos a serem analisados na planilha 'Fontes' " + Chr(10) + Chr(10) + _
'        "- Não enxerga subpastas " + Chr(10) + Chr(10) + _

'        "--------------------------------------------------------------------------------" + Chr(10) + Chr(10) + _
'        "- O passo 1 procura o texto nos fontes (qualquer tipo de texto) " + Chr(10) + _
'        "- O passo 2 procura nos fontes, quais são as funções que chamam as funções do PASSO 1 " + Chr(10) + Chr(10) + _
'        "--------------------------------------------------------------------------------" + Chr(10) + Chr(10) +

MsgBox ("- Os resultados do passo 1 ficam nesta tela e no resumo  " + Chr(10) + _
        "- Os resultados do passo 2 ficam na tela de resumo ")

End Sub

Sub NaoFaz()

MsgBox ("- Não procura em linhas comentadas com // , -- , e */ /* " + Chr(10) + Chr(10) + _
        "- O sistema não entende estático (STATIC FUNCTION), ou seja, acaba procurando em todos os arquivos fonte")
        
End Sub


Sub ChamarForm()
    UserForm1.Show
End Sub




