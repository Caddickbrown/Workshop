Sub ScheduleSort()

    Dim Password As String
    Dim ws As Worksheet
    Password = "bvibutty"
    ws = Worksheets("BVI Main")

    ws.Unprotect Password:=Password 'Unprotect the Sheet with the password

    Rows("1:1048576").EntireRow.Hidden = False 'Unhide any rows
    
    'Clear Filters
    If ActiveSheet.FilterMode = True Then
        ActiveSheet.ShowAllData
    End If
    
    'Sort on Picks
    ws.ListObjects("Table2").Sort.SortFields.Clear
    ws.ListObjects("Table2").Sort.SortFields.Add2 _
        Key:=Range("Table2[[#All],[Picks]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ws.ListObjects("Table2").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Sort on Sequence
    ws.ListObjects("Table2").Sort.SortFields.Clear
    ws.ListObjects("Table2").Sort.SortFields.Add2 _
        Key:=Range("Table2[[#All],[Sequence]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ws.ListObjects("Table2").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Sort on Date
    ws.ListObjects("Table2").Sort.SortFields.Clear
    ws.ListObjects("Table2").Sort.SortFields.Add2 _
        Key:=Range("Table2[[#All],[Date]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ws.ListObjects("Table2").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ws.Protect Password:=Password, AllowSorting:=True, AllowFiltering:=True 'Protect the sheet with the password, allowing filtering

End Sub

Sub MalosaScheduleSort()

    Dim Password As String
    Dim ws As Worksheet
    Password = "malosabutty"
    ws = Worksheets("Malosa Main")

    Worksheets("Malosa Main").Unprotect Password:=Password 'Unprotect the Sheet with the password

    Rows("1:1048576").EntireRow.Hidden = False 'Unhide any rows

    'Clear Filters
    If ActiveSheet.FilterMode = True Then
        ActiveSheet.ShowAllData
    End If
    
    'Sort on Picks
    ws.ListObjects("Table6").Sort.SortFields.Clear
    ws.ListObjects("Table6").Sort.SortFields.Add2 _
        Key:=Range("Table6[[#All],[Picks]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ws.ListObjects("Table6").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'Sort on Sequence
    ws.ListObjects("Table6").Sort.SortFields.Clear
    ws.ListObjects("Table6").Sort.SortFields.Add2 _
        Key:=Range("Table6[[#All],[Sequence]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ws.ListObjects("Table6").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'Sort on Date
    ws.ListObjects("Table6").Sort.SortFields.Clear
    ws.ListObjects("Table6").Sort.SortFields.Add2 _
        Key:=Range("Table6[[#All],[Date]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ws.ListObjects("Table6").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ws.Protect Password:=Password, AllowSorting:=True, AllowFiltering:=True 'Protect the sheet with the password, allowing filtering

End Sub

Sub SampleScheduleSort()

    Dim Password As String
    Dim ws As Worksheet
    Password = "samplesbutty"
    ws = Worksheets("Samples Main")

    ws.Unprotect Password:=Password 'Unprotect the Sheet with the password

    Rows("1:1048576").EntireRow.Hidden = False 'Unhide any rows
    
    'Clear Filters
    If ActiveSheet.FilterMode = True Then
        ActiveSheet.ShowAllData
    End If
    
    'Sort on Picks
    ws.ListObjects("Table29").Sort.SortFields.Clear
    ws.ListObjects("Table29").Sort.SortFields.Add2 _
        Key:=Range("Table29[[#All],[Picks]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ws.ListObjects("Table29").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Sort on Priority
    ws.ListObjects("Table29").Sort.SortFields.Clear
    ws.ListObjects("Table29").Sort.SortFields.Add2 _
        Key:=Range("Table29[[#All],[Priority]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ws.ListObjects("Table29").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Sort on Deadline Completion Date
    ws.ListObjects("Table29").Sort.SortFields.Clear
    ws.ListObjects("Table29").Sort.SortFields.Add2 _
        Key:=Range("Table29[[#All],[Deadline Completion Date]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ws.ListObjects("Table29").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ws.Protect Password:=Password, AllowSorting:=True, AllowFiltering:=True 'Protect the sheet with the password, allowing filtering

End Sub

Sub ArchiveCompleted()
    Dim wsBVI As Worksheet
    Dim wsMalosa As Worksheet
    Dim wsComplete As Worksheet
    Dim tblBVI As ListObject
    Dim tblMalosa As ListObject
    Dim LastRow As Long
    Dim i As Long
    Dim BVIPassword As String
    Dim MalosaPassword As String
    ' Set the password for protecting and unprotecting sheets
    BVIPassword = "bvibutty"
    MalosaPassword = "malosabutty"
    CompletePassword = "completebutty"
    ' Define the destination worksheet as "Complete"
    Set wsComplete = ThisWorkbook.Sheets("Complete") ' Change "Complete" to the name of your destination sheet
    ' Unprotect the destination sheet
    wsComplete.Unprotect Password:=CompletePassword
    
    ' Set the source worksheets based on the provided names
    On Error Resume Next
    Set wsBVI = ThisWorkbook.Sheets("BVI Main")
    Set wsMalosa = ThisWorkbook.Sheets("Malosa Main")
    Set wsSamples = ThisWorkbook.Sheets("Samples Main")
    Set wsComplete = ThisWorkbook.Sheets("Complete")
    Set wsSamplesComplete = ThisWorkbook.Sheets("Samples Complete")
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
    
    'ws.BVI.Select
    'ScheduleSort 'Sort the Schedule into the correct order
    'ws.Malosa.Select
    'MalosaScheduleSort 'Sort the Malosa Schedule into the correct order
    'SampleScheduleSort 'Sort the Samples Schedule into the correct order

    ' Unprotect the source sheets
    wsBVI.Unprotect Password:=BVIPassword
    wsMalosa.Unprotect Password:=MalosaPassword
    
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
    wsBVI.Protect Password:=BVIPassword, AllowSorting:=True, AllowFiltering:=True
    wsMalosa.Protect Password:=MalosaPassword, AllowSorting:=True, AllowFiltering:=True
    wsComplete.Protect Password:=CompletePassword, AllowSorting:=True, AllowFiltering:=True
End Sub


