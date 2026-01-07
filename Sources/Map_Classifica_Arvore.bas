Sub ClassifArvore2()
'
' ClassifArvore2 Macro
'

'
    Columns("R:Z").Select
    ActiveWorkbook.Worksheets("Arvore").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Arvore").Sort.SortFields.Add2 Key:=Range("S2:S501" _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Arvore").Sort.SortFields.Add2 Key:=Range("R2:R501" _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Arvore").Sort
        .SetRange Range("R1:Z501")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWindow.SmallScroll Down:=-39
    Range("T5").Select
End Sub

