Sub ordenafontes()
    
    With ActiveWorkbook.Worksheets("Fontes").Sort
        .SetRange Range("A2:A65536")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

