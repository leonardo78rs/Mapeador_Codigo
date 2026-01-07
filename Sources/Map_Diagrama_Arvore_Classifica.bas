Sub ClassifArvore()
'
' ClassifArvore Macro
'

'
    Sheets("Arvore").Select
    Columns("L:P").Select
    ActiveWorkbook.Worksheets("Arvore").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Arvore").Sort.SortFields.Add2 Key:=Range("M1:M501" _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Arvore").Sort.SortFields.Add2 Key:=Range("L1:L501" _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Arvore").Sort.SortFields.Add2 Key:=Range("N1:N501" _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Arvore").Sort
        .SetRange Range("L1:N501")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With



End Sub

