Sub busca_geral()

Dim pasta As String
Dim Linha As Integer
Dim x As Integer
Dim limite As Integer
Dim cExtensoes As String
Dim cIgnorePastas As String
Dim Fontes As Worksheet

Set Fontes = Sheets("Fontes")

cExtensoes = Fontes.Cells(7, 10)

'Pega o caminho completo da pasta
pasta = InputBox("Confirme o caminho da pasta", "LISTA ARQUIVOS", Fontes.Cells(5, 6))


If Trim(pasta) <> "" Then

    cExtensoes = InputBox("Confirme extensões a mapear", "EXTENSÕES VÁLIDAS", Fontes.Cells(7, 10))
    cIgnorePastas = InputBox("Não procurar nas pastas desta lista... ", "IGNORAR PASTAS", Fontes.Cells(10, 10))
    
    If Right(pasta, 1) <> "\" Then pasta = pasta + "\"
    If Trim(cExtensoes) <> "" Then Fontes.Cells(7, 10) = cExtensoes
    If Trim(cIgnorePastas) <> "" Then Fontes.Cells(10, 10) = cIgnorePastas
    
    Fontes.Cells(5, 6) = pasta
    
    Fontes.Activate
    Range("A2:B65536").Select
    Selection.ClearContents
    
    Linha = lista_arquivos(pasta, 1)
    
    limite = Fontes.Cells(7, 8)
    
    x = 1
    
    Do While Fontes.Cells(x, 1) <> ""
        x = x + 1
        Linha = Fontes.Cells(7, 8)
        
        If Fontes.Cells(x, 2) = "<dir>" Then
            Fontes.Cells(x, 2) = "[dir]"
           Linha = lista_arquivos(Fontes.Cells(x, 1) + "\", Linha)
        End If
    
        
    Loop
    
    x = 1
    Do While Fontes.Cells(x + 1, 1) <> ""
       
       x = x + 1
       If InStr(1, cExtensoes, Right(Fontes.Cells(x, 1), 3), vbTextCompare) = 0 Or _
           Fontes.Cells(x, 2) = "[dir]" Then
        
          Fontes.Cells(x, 1) = ""
          Fontes.Cells(x, 2) = ""
       End If
    
    Loop
    
    
    ordenafontes

End If

Fontes.Cells(3, 8) = Now()

End Sub
