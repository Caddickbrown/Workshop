Sub ProcessMarkedRows()
    ' Declare variables
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim tbl As ListObject
    Dim lastRow As Long, i As Long
    Dim checkCol As Long, markedCount As Long
    Dim calculationResult As Double
    
    ' Set references to worksheets
    Set ws1 = ThisWorkbook.Sheets("Main")
    Set ws2 = ThisWorkbook.Sheets("Line Staffing Options")
    
    ' Reference Table1 (which starts on row 7)
    Set tbl = ws1.ListObjects("Table1")
    
    ' Note: The macro will work correctly with tables regardless of their starting row
    ' Excel tables track their own positions and the ListObject references the correct range
    
    ' Find the Check column in Table1
    checkCol = GetColumnIndex(tbl, "Check")
    
    If checkCol = 0 Then
        MsgBox "Check column not found in Table1", vbExclamation
        Exit Sub
    End If
    
    ' Sort Table1 by Check column (smallest to largest)
    With tbl.Sort
        .SortFields.Clear
        .SortFields.Add Key:=tbl.ListColumns(checkCol).Range, SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With
    
    ' Count rows with "x" in Check column
    markedCount = 0
    For i = 1 To tbl.ListRows.Count
        If tbl.ListRows(i).Range.Cells(1, checkCol).Value = "x" Then
            markedCount = markedCount + 1
        End If
    Next i
    
    ' Exit if no rows marked with "x"
    If markedCount = 0 Then
        MsgBox "No rows with 'x' found in Check column", vbInformation
        Exit Sub
    End If
    
    ' Application.ScreenUpdating = False
    
    For i = 1 To markedCount
        If tbl.ListRows(i).Range.Cells(1, checkCol).Value = "x" Then
            ' Get row index in the worksheet
            Dim actualRow As Long
            actualRow = tbl.ListRows(i).Range.Row
            
            ' Copy values from columns B, F, and X to the destination sheet
            ws2.Cells(1, 16).Value = ws1.Cells(actualRow, 2).Value  ' Column B
            ws2.Cells(1, 18).Value = ws1.Cells(actualRow, 6).Value  ' Column F
            ws2.Cells(1, 21).Value = ws1.Cells(actualRow, 24).Value ' Column X
            
            Calculate
            DoEvents
            
            Application.Wait Now + TimeValue("00:00:02")
            
            ' Copy result back to Table1 (to column Y in this example - adjust as needed)
            ws1.Cells(actualRow, 25).Value = ws2.Cells(1, 23).Value
            ws1.Cells(actualRow, 26).Value = ws2.Cells(1, 24).Value
            
        End If
    Next i
    
    ' Application.ScreenUpdating = True
    
    ' MsgBox "Processing complete. " & markedCount & " rows were processed.", vbInformation
    
End Sub

' Helper function to get column index in a table by name
Function GetColumnIndex(tbl As ListObject, colName As String) As Long
    Dim col As ListColumn
    
    GetColumnIndex = 0
    
    For Each col In tbl.ListColumns
        If col.Name = colName Then
            GetColumnIndex = col.Index
            Exit Function
        End If
    Next col
End Function



