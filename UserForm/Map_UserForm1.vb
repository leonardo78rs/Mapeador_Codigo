'USERFORM 
Private Sub CheckBox1_Click()

Dim wOcorr As Worksheet
Dim marcado As Boolean

Set wOcorr = Sheets("Ocorrencias")

wOcorr.Cells(8, 4) = Not CheckBox1.Value


End Sub


Private Sub CommandButton1_Click()

Principal
ListBox1_Enter

End Sub


Private Sub CommandButton2_Click()
 
Dim wFontes As Worksheet
Set wFontes = Sheets("Fontes")

busca_geral
Label9.Caption = "Quantidade de arquivos: " + Str(wFontes.Cells(6, 8))
Label10.Caption = "Lista de arquivos atualizada em: " + Str(wFontes.Cells(3, 8))

End Sub

Private Sub CommandButton3_Click()
Dim yesnocancel As Integer

yesnocancel = MsgBox("Ação pode ser demorada, deseja remontar árvore?" + Chr(13) + Chr(13) + _
                     " Sim - Remonta " + Chr(13) + _
                     " Não - Apenas visualiza " + Chr(13) + _
                     " Cancela - Retorna" + Chr(13), _
                     vbyesnocancel, _
                      "Alerta")
                     
If yesnocancel = vbYes Then
   Secundario
   UserForm2.Show
ElseIf yesnocancel = vbNo Then
   UserForm2.Show
End If

End Sub

Private Sub CommandButton4_Click()

Dim xquant As Integer
Dim x As Integer
Dim wResumo As Worksheet

Set wResumo = Sheets("Resumo")

If ListBox2.Visible Then

   For x = 1 To ListBox2.ListCount - 2
       wResumo.Cells(x, 25) = ListBox2.List(x + 1)
   Next x
   
   Sheets("Resumo").Select
   Range("Y1").Select
   Range(Selection, Selection.End(xlDown)).Select
   Selection.Copy
   'Selection.Delete

Else ' Caso ListBox2 não está visivel, está copiando a primeira tela ListBox1
    
    Sheets("Resumo").Select
    Range("B3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
End If
End Sub

Private Sub CommandButton5_Click()

If ListBox2.Visible Then
   ListBox2.Visible = False
   CommandButton5.Caption = " Exibir só os fontes "
   CommandButton4.Caption = " Copiar funções <Ctrl><C> "
Else
   ListBox2.Visible = True
   CommandButton5.Caption = " Exibir funções & fontes "
   CommandButton4.Caption = " Copiar fontes <Ctrl><C> "
End If



End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
' Dim x As String
Dim y As Integer
Dim clicado As String
Dim qtdeocorrencias As Integer
Dim funcao As String
Dim fonte As String
Dim barra As Integer
Dim wOcorr As Worksheet
Dim x As Integer
Dim mensagem As String
Dim qtde As Integer
Dim z As Integer


Set wOcorr = Sheets("Ocorrencias")

y = ListBox1.ListIndex
qtde = ListBox1.ListCount

clicado = ListBox1.List(y, 1)

qtdeocorrencias = Val(Mid(clicado, 2, 3))

UserForm3.ListBox1.ColumnCount = 5
UserForm3.ListBox1.ColumnWidths = "0;120;40;40;600"
z = 0

If y < qtde - 3 Then 'descartar 3 linhas finais

   funcao = Mid(clicado, 7, 40)
   x = 13
              
   UserForm3.Caption = Mid(wOcorr.Cells(x, 4), 1, 30)
   
   Do While wOcorr.Cells(x, 1) <> ""
      If Trim(wOcorr.Cells(x, 4)) = Trim(funcao) Then
         'mensagem = mensagem + wOcorr.Cells(x, 5) + Chr(13) + Chr(13)
         'nbarra = InStrRev(ctemp, "\")
         UserForm3.ListBox1.AddItem
         fonte = wOcorr.Cells(x, 1)
         barra = InStrRev(fonte, "\")
         UserForm3.ListBox1.List(z, 1) = Mid(fonte, barra + 1, Len(fonte) - barra)
         UserForm3.ListBox1.List(z, 2) = Str(wOcorr.Cells(x, 2))
         UserForm3.ListBox1.List(z, 3) = Str(wOcorr.Cells(x, 3))
         ' UserForm3.ListBox1.List(z, 4) = Mid(wOcorr.Cells(x, 4), 1, 30)
         UserForm3.ListBox1.List(z, 4) = Trim(wOcorr.Cells(x, 5))
         z = z + 1
      End If
      x = x + 1
   Loop
    
'   MsgBox (mensagem)
End If

'UserForm3.ListBox1.AddItem (cMensagem)
UserForm3.Show


End Sub

Private Sub ListBox1_Enter()

Dim wOcorr As Worksheet
Dim wResumo As Worksheet
Dim lin1, lin2, col As Integer
Dim nini, nfim, nCountGroup As Double
Dim cFonte, cFunct As String
Dim cMensagem As String
Dim cMensagem1 As String
Dim x As Integer
Dim n As Integer
Dim nSumFunc As Integer
Dim nSumOcor As Integer
Dim nPagina As Integer
Dim ctemp As String
Dim nbarra As Integer
Dim nponto As Integer
Dim ultimarq As String
Dim ncontafontes As Integer

Set wOcorr = Sheets("Ocorrencias")
Set wResumo = Sheets("Resumo")
nCountGroup = 0
ListBox1.ColumnCount = 3
ListBox1.ColumnWidths = "0;270"

    x = 3
    nSumOcor = 0
    nPagina = 0
   
    ListBox1.Clear
    ListBox2.Clear
    
    Label6.Caption = "Última busca em: " + Str(wResumo.Cells(1, 1))
        
    Do While wResumo.Cells(x, 1) <> ""
    
       nponto = 0
       nbarra = 0
       
       If wResumo.Cells(x, 2) <> "" Then
          cMensagem = "[" + Trim(Str(wResumo.Cells(x, 3))) + " x]" + wResumo.Cells(x, 2)
       Else
          cMensagem = "[" + Trim(Str(wResumo.Cells(x, 3))) + " x]" + " Sem Função"
       End If
       
       ctemp = wResumo.Cells(x, 1)
       nbarra = InStrRev(ctemp, "\")
       ' nponto = InStrRev(ctemp, ".")
       If nponto = 0 Then nponto = Len(ctemp) - nbarra Else nponto = nponto - nbarra - 1
       
       ctemp = Mid(ctemp, nbarra + 1, nponto)
       
       ' ListBox1.AddItem (cMensagem + Space(50 - Len(cMensagem)) + ctemp)
       ListBox1.AddItem
       ListBox1.List(x - 3, 1) = cMensagem
       ListBox1.List(x - 3, 2) = Trim(ctemp)
       nSumOcor = nSumOcor + wResumo.Cells(x, 3)
              
       x = x + 1
                 
    Loop
    
    '------ Arquivos fonte
    
    If x = 3 Then
       ListBox1.ColumnCount = 2
       cMensagem = " Nenhuma ocorrência "
       ListBox1.AddItem
       ListBox1.List(x - 3, 1) = cMensagem
    End If
    
    nSumFunc = x - 3
         
If nSumFunc > 1 Then
   ListBox1.AddItem
   ListBox1.AddItem
   ListBox1.AddItem
   ListBox1.List(x - 3, 1) = " "
   ListBox1.List(x - 2, 1) = "[" + Str(nSumFunc) + "]  Funções "
   ListBox1.List(x - 1, 1) = "[" + Str(nSumOcor) + "]  Ocorrências "

   '--- tela dos fontes
   ultimarq = ""
   ncontafontes = 0
   ListBox2.AddItem (" Arquivos pesquisados: ")
   ListBox2.AddItem ("")
   For y = 0 To ListBox1.ListCount - 1
       If ultimarq <> ListBox1.List(y, 2) Then
          ultimarq = ListBox1.List(y, 2)
          ListBox2.AddItem (ultimarq)
          ncontafontes = ncontafontes + 1
       End If
   
   Next y
   
   ListBox2.AddItem ("")
   ListBox2.AddItem (Str(ncontafontes) + " Total de arquivos ")
   ' fim tela dos fontes (para ativa-la utilizar outro botao que a deixa visivel)


End If


             
End Sub



Private Sub SpinButton1_SpinUp()

If Label1.Font.Size < 10 Then

    Label1.Font.Size = Label1.Font.Size + 1
    Label2.Font.Size = Label2.Font.Size + 1
    Label3.Font.Size = Label3.Font.Size + 1
    Label4.Font.Size = Label4.Font.Size + 1
    Label9.Font.Size = Label9.Font.Size + 1
    Label10.Font.Size = Label10.Font.Size + 1
    TextBox1.Font.Size = TextBox1.Font.Size + 1
    TextBox2.Font.Size = TextBox2.Font.Size + 1
    TextBox3.Font.Size = TextBox3.Font.Size + 1
    TextBox4.Font.Size = TextBox4.Font.Size + 1
    ListBox1.Font.Size = ListBox1.Font.Size + 1

End If

ListBox1.Height = 200

End Sub

Private Sub SpinButton1_SpinDown()

If Label1.Font.Size > 7 Then
    Label1.Font.Size = Label1.Font.Size - 1
    Label2.Font.Size = Label2.Font.Size - 1
    Label3.Font.Size = Label3.Font.Size - 1
    Label4.Font.Size = Label4.Font.Size - 1
    Label9.Font.Size = Label9.Font.Size - 1
    Label10.Font.Size = Label10.Font.Size - 1
    TextBox1.Font.Size = TextBox1.Font.Size - 1
    TextBox2.Font.Size = TextBox2.Font.Size - 1
    TextBox3.Font.Size = TextBox3.Font.Size - 1
    TextBox4.Font.Size = TextBox4.Font.Size - 1
    ListBox1.Font.Size = ListBox1.Font.Size - 1

End If

ListBox1.Height = 200

End Sub

Private Sub TextBox1_Change()

Dim palavra As String
Dim wOcorr As Worksheet
Set wOcorr = Sheets("Ocorrencias")

wOcorr.Cells(2, 2) = TextBox1.Text

End Sub



Private Sub TextBox2_Change()

Dim pasta As String
Dim wFontes As Worksheet
Set wFontes = Sheets("Fontes")

wFontes.Cells(5, 6) = TextBox2.Text

End Sub


Private Sub TextBox3_Change()

Dim pasta As String
Dim wFontes As Worksheet
Set wFontes = Sheets("Fontes")

wFontes.Cells(7, 10) = TextBox3.Text

End Sub


Private Sub TextBox4_Change()

Dim pasta As String
Dim wFontes As Worksheet
Set wFontes = Sheets("Fontes")

wFontes.Cells(10, 10) = TextBox4.Text
 
End Sub


Private Sub UserForm_Initialize()

Dim palavra As String
Dim pasta As String
Dim wOcorr As Worksheet
Dim wFontes As Worksheet
Set wOcorr = Sheets("Ocorrencias")
Set wFontes = Sheets("Fontes")

TextBox1.Text = wOcorr.Cells(2, 2)
TextBox2.Text = wFontes.Cells(5, 6)
TextBox3.Text = wFontes.Cells(7, 10)
TextBox4.Text = wFontes.Cells(10, 10)
ListBox1_Enter

Label9.Caption = "Quantidade de arquivos: " + Str(wFontes.Cells(6, 8))
Label10.Caption = "Lista de arquivos atualizada em: " + Str(wFontes.Cells(3, 8))

CheckBox1.Value = Not wOcorr.Cells(8, 4)



End Sub



Private Sub ListBox3331_Enter()

Dim wOcorr As Worksheet
Dim wResumo As Worksheet
Dim lin1, lin2, col As Integer
Dim nini, nfim, nCountGroup As Double
Dim cFonte, cFunct As String
Dim cMensagem As String
Dim cMensagem1 As String
Dim x As Integer
Dim n As Integer
Dim nSumFunc As Integer
Dim nSumOcor As Integer
Dim nPagina As Integer

Set wOcorr = Sheets("Ocorrencias")
Set wResumo = Sheets("Resumo")
nCountGroup = 0

    x = 3
    nSumOcor = 0
    nPagina = 0
    
    ListBox1.Clear
    
    Label6.Caption = "Última busca em: " + Str(wResumo.Cells(1, 1))
        
    Do While wResumo.Cells(x, 1) <> ""
    
       If wResumo.Cells(x, 2) <> "" Then
          cMensagem = "   [" + Str(wResumo.Cells(x, 3)) + " x ]" + wResumo.Cells(x, 2)
       Else
          cMensagem = "   [" + Str(wResumo.Cells(x, 3)) + " x ]" + " Sem Função"
       End If
       
       ListBox1.AddItem (cMensagem)
       nSumOcor = nSumOcor + wResumo.Cells(x, 3)
       
       x = x + 1
                 
    Loop
    
    If x = 3 Then
       cMensagem = " Nenhuma ocorrência "
       ListBox1.AddItem (cMensagem)
    End If
    
    nSumFunc = x - 3
         
If nSumFunc > 1 Then
   ListBox1.AddItem ("")
   ListBox1.AddItem (" Total Funções: [" + Str(nSumFunc) + "]  -  Total Ocorrências: [" + Str(nSumOcor) + "]")
End If
             
End Sub


---------------
