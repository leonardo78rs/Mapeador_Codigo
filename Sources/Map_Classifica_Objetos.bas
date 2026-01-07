Sub ClassificaObjetos()
'
' ClassificaObjetos Macro
'

'
    Sheets("Arvore").Select
    Columns("R:AA").Select
    ActiveWorkbook.Worksheets("Arvore").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Arvore").Sort.SortFields.Add2 Key:=Range("r2:r501" _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Arvore").Sort.SortFields.Add2 Key:=Range("s2:s501" _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Arvore").Sort
        .SetRange Range("R1:AA501")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
  
