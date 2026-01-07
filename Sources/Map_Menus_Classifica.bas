Sub ClassifMenuSiac()
'
' ClassifMenuSiac Macro
'

'
    Range("B7:D7").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Menus").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Menus").Sort.SortFields.Add2 Key:=Range("B8:B490") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Menus").Sort.SortFields.Add2 Key:=Range("C8:C490") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Menus").Sort
        .SetRange Range("B7:D490")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub


