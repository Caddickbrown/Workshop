Sub DeleteRows()
' This deletes rows where the Component Requirements are 0 or lower (Column D is for Component Requirement) - basically, where we have enough material

    Dim i As Long
    For i = Range("D" & Rows.Count).End(xlUp).Row To 1 Step -1
        If Not (Range("D" & i).Value > 0) Then
            Range("D" & i).EntireRow.Delete
        End If
    Next i

End Sub


Sub DeleteRows()
    'get last row in column A
    Last = Cells(Rows.Count, "D").End(xlUp).Row
    For i = Last To 1 Step -1
        'if cell value is less than 100
        If (Cells(i, "D").Value) < 100 Then
            'delete entire row
            Cells(i, "D").EntireRow.Delete
        End If
    Next i
End Sub


With Sheets(1).UsedRange
    For lrow = .Rows.Count To 2 Step -1
        If .Cells(lrow, 4).Value <> 110 Then .Rows(lrow).Delete
    Next lrow
End With


'Autofilter

Sub Delete_Rows_Based_On_Value()
'Apply a filter to a Range and delete visible rows

Dim ws As Worksheet

  'Set reference to the sheet in the workbook.
  Set ws = ThisWorkbook.Worksheets("Regular Range")
  ws.Activate 'not required but allows user to view sheet if warning message appears
  
  'Clear any existing filters
  On Error Resume Next
    ws.ShowAllData
  On Error GoTo 0

  '1. Apply Filter
  ws.Range("B3:G1000").AutoFilter Field:=4, Criteria1:=""
  
  '2. Delete Rows
  Application.DisplayAlerts = False
    ws.Range("B4:G1000").SpecialCells(xlCellTypeVisible).Delete
  Application.DisplayAlerts = True
  
  '3. Clear Filter
  On Error Resume Next
    ws.ShowAllData
  On Error GoTo 0

End Sub