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
    ActiveWorkbook.Worksheets("BVI Main").ListObjects("Table2").Sort.SortFields.Clear
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
    ActiveWorkbook.Worksheets("Malosa Main").ListObjects("Table6").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Malosa Main").ListObjects("Table6").Sort.SortFields.Add2 _
        Key:=Range("Table6[[#All],[Picks]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Malosa Main").ListObjects("Table6").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'Sort on Sequence
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
    ActiveWorkbook.Worksheets("Samples Main").ListObjects("Table29").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Samples Main").ListObjects("Table29").Sort.SortFields.Add2 _
        Key:=Range("Table29[[#All],[Picks]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Samples Main").ListObjects("Table29").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Sort on Priority
    ActiveWorkbook.Worksheets("Samples Main").ListObjects("Table29").Sort.SortFields.Clear
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

Sub ArchiveCompleted()
    Dim wsBVI As Worksheet
    Dim wsMalosa As Worksheet
    Dim wsComplete As Worksheet
    Dim tblBVI As ListObject
    Dim tblMalosa As ListObject
    Dim LastRow As Long
    Dim i As Long
    Dim Password As String
    ' Set the password for protecting and unprotecting sheets
    Password = "baconbutty"
    ' Define the destination worksheet as "Complete"
    Set wsComplete = ThisWorkbook.Sheets("Complete") ' Change "Complete" to the name of your destination sheet
    ' Unprotect the destination sheet
    wsComplete.Unprotect Password:=Password
    
    ' Set the source worksheets based on the provided names
    On Error Resume Next
    Set wsBVI = ThisWorkbook.Sheets("BVI Main")
    Set wsMalosa = ThisWorkbook.Sheets("Malosa Main")
    On Error GoTo 0
    
    If wsBVI Is Nothing Or wsMalosa Is Nothing Then
        MsgBox "One or both of the source sheets does not exist."
        Exit Sub
    End If
    
    ' Set the source tables based on the provided names
    On Error Resume Next
    Set tblBVI = wsBVI.ListObjects("Table2") ' Kits
    Set tblMalosa = wsMalosa.ListObjects("Table6") ' Kits
    'Set tblBVI = wsBVI.ListObjects("Table1") ' Instruments
    'Set tblMalosa = wsMalosa.ListObjects("Table15") ' Instruments
    On Error GoTo 0
    
    If tblBVI Is Nothing Or tblMalosa Is Nothing Then
        MsgBox "One or both of the source tables does not exist."
        Exit Sub
    End If
    
    ScheduleSort 'Sort the Schedule into the correct order
    MalosaScheduleSort 'Sort the Malosa Schedule into the correct order
    SampleScheduleSort 'Sort the Samples Schedule into the correct order

    ' Unprotect the source sheets
    wsBVI.Unprotect Password:=Password
    wsMalosa.Unprotect Password:=Password
    
    ' Find the last row in the source tables and move completed orders
    For Each tbl In Array(tblBVI, tblMalosa)
        For i = tbl.ListRows.Count To 1 Step -1
            If tbl.ListRows(i).Range.Cells(1, tbl.ListColumns("Status").Index).Value = "Completed" Then
                ' Copy the entire row to the destination sheet
                tbl.ListRows(i).Range.Copy wsComplete.Cells(wsComplete.Cells(wsComplete.Rows.Count, "A").End(xlUp).Row + 1, 1)
                
                ' Delete the row from the source table (optional)
                tbl.ListRows(i).Delete
            End If
        Next i
    Next tbl
    
    ' Protect the source sheets and the destination sheet again
    wsBVI.Protect Password:=Password, AllowSorting:=True, AllowFiltering:=True
    wsMalosa.Protect Password:=Password, AllowSorting:=True, AllowFiltering:=True
    wsComplete.Protect Password:=Password, AllowSorting:=True, AllowFiltering:=True
End Sub

