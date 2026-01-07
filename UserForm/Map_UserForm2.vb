'USERFORM 

Private Sub CommandButton4_Click()
Dim xlinhas As Integer
Dim xrange As String

xlinhas = ListBox1.ListCount
xrange = "A1:C" + Trim(Str(xlinhas))

Sheets("Arvore").Select
Range(xrange).Select
Selection.Copy

End Sub

Private Sub ListBox1_Enter()

Dim Arvore As Worksheet
Dim cMensagem As String
Dim x As Integer


Set Arvore = Sheets("Arvore")
nCountGroup = 0
cMensagem = ""
x = 1
  
    ListBox1.Clear
    'Arvore.Cells(X, 1) = Str(Arvore.Cells(X, 1))
    
    Do While Arvore.Cells(x, 1) <> "" Or _
             Arvore.Cells(x, 2) <> "" Or _
             Arvore.Cells(x, 3) <> ""

       If IsNumeric(Arvore.Cells(x, 1)) And Trim(Arvore.Cells(x, 1)) <> Empty Then
          cMensagem = Str(Arvore.Cells(x, 1)) + Space(8) + Arvore.Cells(x, 2) + Space(12) + Arvore.Cells(x, 3)
       Else
          cMensagem = Arvore.Cells(x, 1) + Space(8) + Arvore.Cells(x, 2) + Space(12) + Arvore.Cells(x, 3)
       End If
       
       ListBox1.AddItem (cMensagem)
       
     
       
       
       x = x + 1
                 
    Loop
    
    If x = 1 Then
       cMensagem = " Nenhuma ocorrência "
       ListBox1.AddItem (cMensagem)
    End If
    
    Label1.Caption = "Dados obtidos em: " + Str(Arvore.Cells(1, 5))

    
    
End Sub
