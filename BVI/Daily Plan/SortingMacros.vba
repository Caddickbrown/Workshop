'ToDo
'- [ ] Add calling Subs/Abstractions to shorten Sorts (Mostly the same anyway)

Sub ScheduleSort()
'Sorts the Sheet into Schedule Order

    Worksheets("BVI Main").Unprotect Password:="baconbutty" 'Unprotect the Sheet with the password

    Rows("1:1048576").EntireRow.Hidden = False 'Unhide any rows
    
    'Clear Filters
    If ActiveSheet.FilterMode = True Then
        ActiveSheet.ShowAllData
    End If
    
    'Sort on Picks
    ActiveWorkbook.Worksheets("BVI Main").ListObjects("Table2").Sort.SortFields.Add2 _
        Key:=Range("Table2[[#All],[Picks]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BVI Main").ListObjects("Table2").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Sort on Sequence
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
    
    'Sort on Date
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

    ActiveSheet.Protect Password:="baconbutty", AllowSorting:=True, AllowFiltering:=True 'Protect the sheet with the password, allowing filtering

End Sub

Sub MalosaScheduleSort()

    Worksheets("Malosa Main").Unprotect Password:="baconbutty" 'Unprotect the Sheet with the password

    Rows("1:1048576").EntireRow.Hidden = False 'Unhide any rows

    'Clear Filters
    If ActiveSheet.FilterMode = True Then
        ActiveSheet.ShowAllData
    End If
    
    'Sort on Picks
    ActiveWorkbook.Worksheets("BVI Main").ListObjects("Table2").Sort.SortFields.Add2 _
        Key:=Range("Table2[[#All],[Picks]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BVI Main").ListObjects("Table2").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'Sort on Sequence
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

    'Sort on Date
    ActiveWorkbook.Worksheets("Malosa Main").ListObjects("Table6").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Malosa Main").ListObjects("Table6").Sort.SortFields.Add2 _
        Key:=Range("Table6[[#All],[Date]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Malosa Main").ListObjects("Table6").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ActiveSheet.Protect Password:="baconbutty", AllowSorting:=True, AllowFiltering:=True 'Protect the sheet with the password, allowing filtering

End Sub

Sub SampleScheduleSort()
'Sorts the Sheet into Schedule Order

    Worksheets("Samples Main").Unprotect Password:="baconbutty" 'Unprotect the Sheet with the password

    Rows("1:1048576").EntireRow.Hidden = False 'Unhide any rows
    
    'Clear Filters
    If ActiveSheet.FilterMode = True Then
        ActiveSheet.ShowAllData
    End If
    
    'Sort on Picks
    ActiveWorkbook.Worksheets("BVI Main").ListObjects("Table2").Sort.SortFields.Add2 _
        Key:=Range("Table2[[#All],[Picks]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BVI Main").ListObjects("Table2").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Sort on Priority
    ActiveWorkbook.Worksheets("Samples Main").ListObjects("Table29").Sort.SortFields.Add2 _
        Key:=Range("Table29[[#All],[Priority]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Samples Main").ListObjects("Table29").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Sort on Deadline Completion Date
    ActiveWorkbook.Worksheets("Samples Main").ListObjects("Table29").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Samples Main").ListObjects("Table29").Sort.SortFields.Add2 _
        Key:=Range("Table29[[#All],[Deadline Completion Date]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Samples Main").ListObjects("Table29").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ActiveSheet.Protect Password:="baconbutty", AllowSorting:=True, AllowFiltering:=True 'Protect the sheet with the password, allowing filtering

End Sub

