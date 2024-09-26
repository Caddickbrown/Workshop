'Issues Log

Sub ILSort()
    Dim sortColumns As Variant
    sortColumns = [{"Owner", "Date Added"}]
    ScheduleSort Worksheets("Issues Log"), "Table1", sortColumns
End Sub

Sub CompletedSort()
    Dim sortColumns As Variant
    sortColumns = [{"Owner", "Date Added"}]
    ScheduleSort Worksheets("Archive"), "Table3", sortColumns
End Sub

Sub AllScheduleSort()
    ILSort
    CompletedSort
End Sub

Sub ScheduleSort(ws As Worksheet, tableName As String, sortColumns As Variant)
    ws.Select
    
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

End Sub

Sub ArchiveCompleted()
    Dim wsIL As Worksheet, wsComplete As Worksheet
    Dim tblIL As ListObject
    Dim LastRow As Long
    Dim i As Long
    Dim Password As String
    
    ' Set the password for protecting and unprotecting sheets

    ' Define the destination worksheet as "Archive"
    Set wsComplete = ThisWorkbook.Sheets("Archive")
        
    ' Set the source worksheets based on the provided names
    On Error Resume Next
    Set wsIL = ThisWorkbook.Sheets("Issues Log")
    On Error GoTo 0
    
    If wsIL Is Nothing Or wsComplete Is Nothing Then
        MsgBox "One or both of the source sheets does not exist."
        Exit Sub
    End If
    
    AllScheduleSort
    
    ' Set the source tables based on the provided names
    On Error Resume Next
    Set tblIL = wsIL.ListObjects("Table1")
    On Error GoTo 0
    
    If tblIL Is Nothing Then
        MsgBox "The source table does not exist."
        Exit Sub
    End If
    
    Application.Calculation = xlManual
    
    ' Find the last row in the source tables and move completed orders
    For Each tbl In Array(tblIL)
    
        For i = tbl.ListRows.Count To 1 Step -1
            If tbl.ListRows(i).Range.Cells(1, tbl.ListColumns("Closed").Index).Value = "Y" Then
                ' Copy the entire row to the destination sheet
                tbl.ListRows(i).Range.Copy wsComplete.Cells(wsComplete.Cells(wsComplete.Rows.Count, "A").End(xlUp).Row + 1, 1)
                
                ' Delete the row from the source table (optional)
                tbl.ListRows(i).Delete
            End If
        Next i
    Next tbl

    Application.Calculation = xlAutomatic

    wsComplete.Columns("A:K").FormatConditions.Delete

    ' Save
    ActiveWorkbook.Save

End Sub

' # Changelog

' ## [1.0.0] - 2024-09-26

' ### Added

' - Init Commit
' - Changelog