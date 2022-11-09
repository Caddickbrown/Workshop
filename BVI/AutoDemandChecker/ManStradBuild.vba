'ToDo
' - [ ] Update Supplier Lookup
' - [ ] Code Documentation

Sub ManStradBuild()

'Prep
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

'Complete Calculations
    Do
        DoEvents
        Application.Calculate
    Loop While Not Application.CalculationState = xlDone

'Clear out Negatives
    Dim i As Long
    For i = Range("D" & Rows.Count).End(xlUp).Row To 1 Step -1
        If Not (Range("D" & i).Value > 0) Then
            Range("D" & i).EntireRow.Delete
        End If
    Next i
    
'Cleanup
    Cells.EntireColumn.AutoFit
    Range("A1").Select
End Sub