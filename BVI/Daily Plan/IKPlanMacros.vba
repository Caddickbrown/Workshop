Sub ScheduleMSort()
'Sorts the BVI Manufacturing Sheet into Schedule Order

    Worksheets("BVI Manufacturing").Unprotect Password:="baconbutty" 'Unprotect the Sheet with the password

    Worksheets("BVI Manufacturing").Rows("1:1048576").EntireRow.Hidden = False 'Unhide any rows
    
    'Clear Filters
    If Worksheets("BVI Manufacturing").FilterMode = True Then
        Worksheets("BVI Manufacturing").ShowAllData
    End If
    
    'Sort on Sequence
    ActiveWorkbook.Worksheets("BVI Manufacturing").ListObjects("Table19").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BVI Manufacturing").ListObjects("Table19").Sort.SortFields.Add2 _
        Key:=Range("Table19[[#All],[Sequence]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BVI Manufacturing").ListObjects("Table19").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Sort on Date
    ActiveWorkbook.Worksheets("BVI Manufacturing").ListObjects("Table19").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BVI Manufacturing").ListObjects("Table19").Sort.SortFields.Add2 _
        Key:=Range("Table19[[#All],[Date]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BVI Manufacturing").ListObjects("Table19").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Worksheets("BVI Manufacturing").Protect Password:="baconbutty", AllowSorting:=True, AllowFiltering:=True 'Protect the sheet with the password, allowing filtering

End Sub

Sub ScheduleASort()
'Sorts the BVI Assembly Sheet into Schedule Order

    Worksheets("BVI Assembly").Unprotect Password:="baconbutty" 'Unprotect the Sheet with the password

    Worksheets("BVI Assembly").Rows("1:1048576").EntireRow.Hidden = False 'Unhide any rows
    
    'Clear Filters
    If Worksheets("BVI Assembly").FilterMode = True Then
        Worksheets("BVI Assembly").ShowAllData
    End If
    
    'Sort on Sequence
    ActiveWorkbook.Worksheets("BVI Assembly").ListObjects("Table1910").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BVI Assembly").ListObjects("Table1910").Sort.SortFields.Add2 _
        Key:=Range("Table1[[#All],[Sequence]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BVI Assembly").ListObjects("Table1910").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Sort on Date
    ActiveWorkbook.Worksheets("BVI Assembly").ListObjects("Table1910").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BVI Assembly").ListObjects("Table1910").Sort.SortFields.Add2 _
        Key:=Range("Table1[[#All],[Date]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BVI Assembly").ListObjects("Table1910").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Worksheets("BVI Assembly").Protect Password:="baconbutty", AllowSorting:=True, AllowFiltering:=True 'Protect the sheet with the password, allowing filtering

End Sub

Sub SchedulePSort()
'Sorts the BVI Packaging Sheet into Schedule Order

    Worksheets("BVI Packaging").Unprotect Password:="baconbutty" 'Unprotect the Sheet with the password

    Worksheets("BVI Packaging").Rows("1:1048576").EntireRow.Hidden = False 'Unhide any rows
    
    'Clear Filters
    If Worksheets("BVI Packaging").FilterMode = True Then
        Worksheets("BVI Packaging").ShowAllData
    End If
    
    'Sort on Sequence
    ActiveWorkbook.Worksheets("BVI Packaging").ListObjects("Table1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BVI Packaging").ListObjects("Table1").Sort.SortFields.Add2 _
        Key:=Range("Table1[[#All],[Sequence]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BVI Packaging").ListObjects("Table1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Sort on Date
    ActiveWorkbook.Worksheets("BVI Packaging").ListObjects("Table1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BVI Packaging").ListObjects("Table1").Sort.SortFields.Add2 _
        Key:=Range("Table1[[#All],[Date]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BVI Packaging").ListObjects("Table1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Worksheets("BVI Packaging").Protect Password:="baconbutty", AllowSorting:=True, AllowFiltering:=True 'Protect the sheet with the password, allowing filtering

End Sub

Sub MalosaScheduleSort()


    Worksheets("Malosa Main").Unprotect Password:="baconbutty" 'Unprotect the Sheet with the password

    Worksheets("Malosa Main").Rows("1:1048576").EntireRow.Hidden = False 'Unhide any rows

    'Clear Filters
    If Worksheets("Malosa Main").FilterMode = True Then
        Worksheets("Malosa Main").ShowAllData
    End If

    'Sort on Sequence
    ActiveWorkbook.Worksheets("Malosa Main").ListObjects("Table15").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Malosa Main").ListObjects("Table15").Sort.SortFields.Add2 _
        Key:=Range("Table15[[#All],[Sequence]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Malosa Main").ListObjects("Table15").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'Sort on Date
    ActiveWorkbook.Worksheets("Malosa Main").ListObjects("Table15").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Malosa Main").ListObjects("Table15").Sort.SortFields.Add2 _
        Key:=Range("Table15[[#All],[Date]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Malosa Main").ListObjects("Table15").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Worksheets("Malosa Main").Protect Password:="baconbutty", AllowSorting:=True, AllowFiltering:=True 'Protect the sheet with the password, allowing filtering

End Sub

Sub CompletedScheduleSort()

    Worksheets("Complete").Unprotect Password:="baconbutty" 'Unprotect the Sheet with the password

    ThisWorkbook.Worksheets("Complete").Rows("1:1048576").EntireRow.Hidden = False 'Unhide any rows

    'Clear Filters
    If Worksheets("Complete").FilterMode = True Then
        Worksheets("Complete").ShowAllData
    End If

    'Sort on Sequence
    ActiveWorkbook.Worksheets("Complete").ListObjects("Table7").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Complete").ListObjects("Table7").Sort.SortFields.Add2 _
        Key:=Range("Table7[[#All],[Sequence]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Complete").ListObjects("Table7").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'Sort on Date
    ActiveWorkbook.Worksheets("Complete").ListObjects("Table7").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Complete").ListObjects("Table7").Sort.SortFields.Add2 _
        Key:=Range("Table7[[#All],[Date]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Complete").ListObjects("Table7").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Worksheets("Complete").Protect Password:="baconbutty", AllowSorting:=True, AllowFiltering:=True 'Protect the sheet with the password, allowing filtering

End Sub

Sub ArchiveCompleted()
    Dim wsBVIM As Worksheet, wsBVIA As Worksheet, wsBVIP As Worksheet, wsMalosa As Worksheet, wsComplete As Worksheet
    Dim tblBVIM As ListObject, tblBVIA As ListObject, tblBVIP As ListObject, tblMalosa As ListObject
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
    Set wsBVIM = ThisWorkbook.Sheets("BVI Manufacturing")
    Set wsBVIA = ThisWorkbook.Sheets("BVI Assembly")
    Set wsBVIP = ThisWorkbook.Sheets("BVI Packaging")
    Set wsMalosa = ThisWorkbook.Sheets("Malosa Main")
    On Error GoTo 0
    
    If wsBVIM Is Nothing Or wsBVIA Is Nothing Or wsBVIP Is Nothing Or wsMalosa Is Nothing Then
        MsgBox "One or both of the source sheets does not exist."
        Exit Sub
    End If
    
    wsBVIM.Select
    ScheduleMSort
    wsBVIA.Select
    ScheduleASort
    wsBVIP.Select
    SchedulePSort
    wsMalosa.Select
    MalosaScheduleSort
    wsComplete.Select
    
    ' Set the source tables based on the provided names
    On Error Resume Next
    Set tblBVIM = wsBVIM.ListObjects("Table19")
    Set tblBVIA = wsBVIA.ListObjects("Table1910")
    Set tblBVIP = wsBVIP.ListObjects("Table1")
    Set tblMalosa = wsMalosa.ListObjects("Table15")
    On Error GoTo 0
    
    If tblBVIM Is Nothing Or tblBVIA Is Nothing Or tblBVIP Is Nothing Or tblMalosa Is Nothing Then
        MsgBox "One or both of the source tables does not exist."
        Exit Sub
    End If
    
    ScheduleMSort 'Sort the Manufacturing Schedule into the correct order
    ScheduleASort 'Sort the Assembly Schedule into the correct order
    SchedulePSort 'Sort the Packaging Schedule into the correct order
    MalosaScheduleSort 'Sort the Malosa Schedule into the correct order
    
    ' Unprotect the source sheets
    wsBVIM.Unprotect Password:=Password
    wsBVIA.Unprotect Password:=Password
    wsBVIP.Unprotect Password:=Password
    wsMalosa.Unprotect Password:=Password
    wsComplete.Unprotect Password:=Password
    
    ' Find the last row in the source tables and move completed orders
    For Each tbl In Array(tblBVIM, tblBVIA, tblBVIP, tblMalosa)
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
    wsBVIM.Protect Password:=Password, AllowSorting:=True, AllowFiltering:=True
    wsBVIA.Protect Password:=Password, AllowSorting:=True, AllowFiltering:=True
    wsBVIP.Protect Password:=Password, AllowSorting:=True, AllowFiltering:=True
    wsMalosa.Protect Password:=Password, AllowSorting:=True, AllowFiltering:=True
    wsComplete.Protect Password:=Password, AllowSorting:=True, AllowFiltering:=True
End Sub



