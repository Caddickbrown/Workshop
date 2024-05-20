' Todo
' - [ ] Move the sort so that it sorts just before it processes the sheet.

Sub ScheduleMSort()
    Dim sortColumns As Variant
    sortColumns = [{"Sequence", "Date"}]
    ScheduleSort Worksheets("BVI Manufacturing"), "Table19", sortColumns
End Sub

Sub ScheduleASort()
    Dim sortColumns As Variant
    sortColumns = [{"Sequence", "Date"}]
    ScheduleSort Worksheets("BVI Assembly"), "Table1910", sortColumns
End Sub

Sub SchedulePSort()
    Dim sortColumns As Variant
    sortColumns = [{"Sequence", "Date"}]
    ScheduleSort Worksheets("BVI Packaging"), "Table1", sortColumns
End Sub

Sub MalosaScheduleSort()
    Dim sortColumns As Variant
    sortColumns = [{"Sequence", "Date"}]
    ScheduleSort Worksheets("Malosa Main"), "Table15", sortColumns
End Sub

Sub CompletedScheduleSort()
    Dim sortColumns As Variant
    sortColumns = [{"Sequence", "Date"}]
    ScheduleSort Worksheets("Complete"), "Table7", sortColumns
End Sub

Sub ScheduleSort(ws As Worksheet, tableName As String, sortColumns As Variant)
    ws.Select
    Protection ws, "Unprotect"
    
    ' Unhide any rows
    ws.Rows("1:1048576").EntireRow.Hidden = False
    
    ' Clear Filters
    If ws.FilterMode = True Then
        ws.ShowAllData
    End If
    
    ' Loop through each sort column
    For Each sortColumn In sortColumns
        ' Sort on the current column
        ws.ListObjects(tableName).Sort.SortFields.Clear
        ws.ListObjects(tableName).Sort.SortFields.Add2 _
            Key:=Range(tableName & "[[#All],[" & sortColumn & "]]"), SortOn:=xlSortOnValues, Order:= _
            xlAscending, DataOption:=xlSortNormal
        With ws.ListObjects(tableName).Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    Next sortColumn
    
    ' Protect the sheet with the password, allowing sorting and filtering
    Protection ws, "Protect"
End Sub


Sub Protection(obj As Object, action As String)
    Dim Password As String
    Password = "baconbutty"

    Select Case action
        Case "Protect"
            obj.Protect Password:=Password, AllowSorting:=True, AllowFiltering:=True ', UserInterfaceOnly:=True
        Case "Unprotect"
            obj.Unprotect Password:=Password
        Case Else
            ' Throw an error for an invalid action
            Err.Raise vbObjectError + 9999, "Protection", "Invalid action. Use 'Protect' or 'Unprotect'."
    End Select
End Sub

Sub ArchiveCompleted()
    Dim wsBVIM As Worksheet, wsBVIA As Worksheet, wsBVIP As Worksheet, wsMalosa As Worksheet, wsComplete As Worksheet
    Dim tblBVIM As ListObject, tblBVIA As ListObject, tblBVIP As ListObject, tblMalosa As ListObject
    Dim LastRow As Long
    Dim i As Long
    Dim Password As String
    
    ' Set the password for protecting and unprotecting sheets

    ' Define the destination worksheet as "Complete"
    Set wsComplete = ThisWorkbook.Sheets("Complete") ' Change "Complete" to the name of your destination sheet
        
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
    
    ScheduleMSort
    ScheduleASort
    SchedulePSort
    MalosaScheduleSort
    CompletedScheduleSort
    
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
    
    Protection wsComplete, "Unprotect"
    
    ' Find the last row in the source tables and move completed orders
    For Each tbl In Array(tblBVIM, tblBVIA, tblBVIP, tblMalosa)
        Protection tbl.Parent, "Unprotect"
        For i = tbl.ListRows.Count To 1 Step -1
            If tbl.ListRows(i).Range.Cells(1, tbl.ListColumns("Status").Index).Value = "Completed" Then
                ' Copy the entire row to the destination sheet
                tbl.ListRows(i).Range.Copy wsComplete.Cells(wsComplete.Cells(wsComplete.Rows.Count, "A").End(xlUp).Row + 1, 1)
                
                ' Delete the row from the source table (optional)
                tbl.ListRows(i).Delete
            End If
        Next i
        Protection tbl.Parent, "Protect"
    Next tbl

    wsComplete.Columns("A:O").FormatConditions.Delete

    Protection wsComplete, "Protect"
End Sub