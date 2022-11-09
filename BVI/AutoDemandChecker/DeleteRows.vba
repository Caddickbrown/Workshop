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
    ws.Range("B2:G1000").SpecialCells(xlCellTypeVisible).Delete
  Application.DisplayAlerts = True
  
  '3. Clear Filter
  On Error Resume Next
    ws.ShowAllData
  On Error GoTo 0

End Sub




Sub ManStradBuild()

'Prep
    If ActiveSheet.AutoFilterMode Then ActiveSheet.ShowAllData
    Sheets("ManStrad").Cells.ClearContents

'Copy Data Over
    Sheets("ManStrad").Range("A:D").Value = Sheets("ManStructures").Range("A:D").Value
    Columns("B:B").Delete Shift:=xlToLeft

'Headers
    Range("D1:F1").Value = Array("Component Requirement", "Supplier", "Comments")
    Range("G1").Formula = "=TODAY()-WEEKDAY(TODAY(),3)"
    Range("H1:L1").FormulaR1C1 = "=RC[-1]+7"


    lrtarget = ActiveWorkbook.Sheets("ManStrad").Range("A1", Sheets("ManStrad").Range("A1").End(xlDown)).Rows.Count

    Range("D2:D" & lrtarget).FormulaR1C1 = "=SUM(SUMIF('Reqs'!C[-2],RC[-3],'Reqs'!C),SUMIFS('Open Orders'!C[1],'Open Orders'!C[-2],RC[-3],'Open Orders'!C[-1],""Released""))"
    Range("E2:E" & lrtarget).FormulaR1C1 = "=VLOOKUP(RC[-3],'Purchase Order Lines'!C[-3]:C[-2],2,FALSE)"
    Range("G2:G" & lrtarget).FormulaR1C1 = "=SUM(SUMIFS('Reqs'!C4,'Reqs'!C2,RC1,'Reqs'!C3,""<=""&R1C[1]+6),SUMIFS('Open Orders'!C5,'Open Orders'!C2,RC1,'Open Orders'!C4,""<=""&R1C[1]+6,'Open Orders'!C3,""Released""))*RC3"
    Range("H2:L" & lrtarget).FormulaR1C1 = "=SUM(SUMIFS('Reqs'!C4,'Reqs'!C2,RC1,'Reqs'!C3,""<=""&R1C[1]+6,'Reqs'!C3,"">=""&R1C[1]),SUMIFS('Open Orders'!C5,'Open Orders'!C2,RC1,'Open Orders'!C4,""<=""&R1C[1]+6,'Open Orders'!C3,""Released"",'Open Orders'!C4,"">=""&R1C[1]))*RC3"
   
    
    If Range("A1").AutoFilterMode = True Then
    Else
    Range("A1").AutoFilter
    End If
    
    
    
    ActiveSheet.Range("$A:$L" & lrtarget).AutoFilter Field:=4, Criteria1:="<1", Operator:=xlAnd
    Range("$A:$L" & lrtarget).SpecialCells(xlCellTypeVisible).Delete
    
'Cleanup
    Cells.EntireColumn.AutoFit
    Range("A1").Select
End Sub
