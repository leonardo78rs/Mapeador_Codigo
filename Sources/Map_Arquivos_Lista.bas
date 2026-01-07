Function lista_arquivos(pasta As String, Linha As Integer)
    
    Dim arquivo As String
    Dim x As Integer
    Dim nArqIni As Integer
    Dim nArqFim As Integer
    Dim aListaArquivos As Variant
    Dim nLoop As Integer
    Dim nExiste As Integer
    Dim nPrimaArq As Integer
    Dim nPrimaPas As Integer
    Dim cIgnorePastas As String
    Dim Fontes As Worksheet

    Set Fontes = Sheets("Fontes")
    
    cIgnorePastas = Fontes.Cells(10, 10)
 
    'Existe um problema no Dir():
    ' no primeiro coloca os argumentos, a partir do segundo não deve ter argumentos para trazer os proximos
    nPrimaArq = 0
    nPrimaPas = 0
    
    nArqIni = Linha
    
    'Lista os arquivos restantes
    Do While arquivo <> "" Or nPrimaArq = 0
        
        If nPrimaArq = 0 Then
           arquivo = Dir(pasta, vbArchive)
           nPrimaArq = 1
        Else
           arquivo = Dir
        End If
                
        If arquivo <> "" Then
            Linha = Linha + 1
            Fontes.Cells(Linha, 1) = pasta + arquivo
        End If
    Loop
    Fontes.Cells(7, 8) = Linha
    
    nArqFim = Linha
   
    'Lista as pastas
    Do While arquivo <> "" Or nPrimaPas = 0
    
        If nPrimaPas = 0 Then
           arquivo = Dir(pasta, vbDirectory)
           nPrimaPas = 1
        Else
           arquivo = Dir
        End If
        
        If arquivo <> "" And _
           arquivo <> "." And _
           arquivo <> ".." Then
            
            ' o comando vbdirectory traz arquivos e pastas (por isto este bloco - para considerar so as pastas)
            nExiste = 0
            
            For nLoop = nArqIni To nArqFim
                    
                If Fontes.Cells(nLoop, 1) = pasta + arquivo Then
                   nExiste = nExiste + 1
                End If
                
            Next
            
            ' se nao existe é porque não é arquivo, é pasta
            ' ver também se o nome da pasta é para ignorar
            If nExiste = 0 And _
               InStr(1, cIgnorePastas, arquivo, vbTextCompare) = 0 Then
               
               Linha = Linha + 1
               Fontes.Cells(Linha, 1) = pasta + arquivo
               Fontes.Cells(Linha, 2) = "<dir>"
               
            End If
            
        End If
    Loop
    Fontes.Cells(7, 8) = Linha
    
    ' Fim pastas
    
    Do While Fontes.Cells(Linha + 1, 1) <> ""
            Linha = Linha + 1
            Fontes.Cells(Linha, 1) = ""
    Loop
    
End Function

