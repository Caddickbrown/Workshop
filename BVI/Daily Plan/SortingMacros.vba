Sub ScheduleSort()
' Sorts the Sheet into Schedule Order

    Worksheets("BVI Main").Unprotect Password:="baconbutty"

    ActiveWorkbook.Worksheets("BVI Main").ListObjects("Table2").Sort.SortFields.Clear
    Rows("1:1048576").EntireRow.Hidden = False
    ActiveWorkbook.Worksheets("BVI Main").ListObjects("Table2").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BVI Main").ListObjects("Table2").Sort.SortFields.Add2 _
        Key:=Range("Table2[[#All],[Sequence]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BVI Main").ListObjects("Table2").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ActiveWorkbook.Worksheets("BVI Main").ListObjects("Table2").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BVI Main").ListObjects("Table2").Sort.SortFields.Add2 _
        Key:=Range("Table2[[#All],[Date]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BVI Main").ListObjects("Table2").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Worksheets("BVI Main").Protect Password:="baconbutty"

End Sub

Sub MalosaScheduleSort()

    Worksheets("Malosa Main").Unprotect Password:="baconbutty"

    ActiveWorkbook.Worksheets("Malosa Main").ListObjects("Table6").Sort.SortFields.Clear
    Rows("1:1048576").EntireRow.Hidden = False
    ActiveWorkbook.Worksheets("Malosa Main").ListObjects("Table6").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Malosa Main").ListObjects("Table6").Sort.SortFields.Add2 _
        Key:=Range("Table6[[#All],[Sequence]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Malosa Main").ListObjects("Table6").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ActiveWorkbook.Worksheets("Malosa Main").ListObjects("Table6").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Malosa Main").ListObjects("Table6").Sort.SortFields.Add2 _
        Key:=Range("Table6[[#All],[Ship No.]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Malosa Main").ListObjects("Table6").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Worksheets("Malosa Main").Protect Password:="baconbutty"

End Sub
