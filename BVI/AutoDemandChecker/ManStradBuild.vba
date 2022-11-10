'ToDo
' - [ ] Update Supplier Lookup
' - [ ] Code Documentation

Sub ManStradBuild()

'Prep
    If (ActiveSheet.AutoFilterMode And ActiveSheet.FilterMode) Or ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
    Sheets("ManStrad").Cells.ClearContents

'Copy Data Over
    Sheets("ManStrad").Range("A:D").Value = Sheets("ManStructures").Range("A:D").Value
    Columns("B:B").Delete Shift:=xlToLeft

'Headers
    Range("D1:F1").Value = Array("Component Requirement", "Supplier", "Comments")
    Range("G1").Formula = "=TODAY()-WEEKDAY(TODAY(),3)"
    Range("H1:N1").FormulaR1C1 = "=RC[-1]+7"


    lrtarget = ActiveWorkbook.Sheets("ManStrad").Range("A1", Sheets("ManStrad").Range("A1").End(xlDown)).Rows.Count

    Range("D2:D" & lrtarget).FormulaR1C1 = "=SUM(SUMIF('Reqs'!C[-2],RC[-3],'Reqs'!C),SUMIFS('Open Orders'!C[1],'Open Orders'!C[-2],RC[-3],'Open Orders'!C[-1],""Released""))"
    Range("E2:E" & lrtarget).FormulaR1C1 = "=VLOOKUP(RC[-3],'Suppliers'!C[-4]:C[-3],2,FALSE)"
    Range("G2:G" & lrtarget).FormulaR1C1 = "=SUM(SUMIFS('Reqs'!C4,'Reqs'!C2,RC1,'Reqs'!C3,""<=""&R1C+6),SUMIFS('Open Orders'!C5,'Open Orders'!C2,RC1,'Open Orders'!C4,""<=""&R1C+6,'Open Orders'!C3,""Released""))*RC3"
    Range("H2:N" & lrtarget).FormulaR1C1 = "=SUM(SUMIFS('Reqs'!C4,'Reqs'!C2,RC1,'Reqs'!C3,""<=""&R1C+6,'Reqs'!C3,"">=""&R1C),SUMIFS('Open Orders'!C5,'Open Orders'!C2,RC1,'Open Orders'!C4,""<=""&R1C+6,'Open Orders'!C3,""Released"",'Open Orders'!C4,"">=""&R1C))*RC3"
   
    Range("D1").AutoFilter
    ActiveWorkbook.Worksheets("ManStrad").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("ManStrad").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("D:D"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("ManStrad").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ActiveSheet.Range("$A1:$N" & lrtarget).AutoFilter Field:=4, Criteria1:="<1", Operator:=xlAnd
    
    Application.DisplayAlerts = False
    Range("$A2:$N" & lrtarget).SpecialCells(xlCellTypeVisible).Delete
    Application.DisplayAlerts = True
    
'Cleanup
    ActiveSheet.Range("$A:$N").AutoFilter Field:=4
    Cells.EntireColumn.AutoFit
    Range("A1").Select

End Sub

