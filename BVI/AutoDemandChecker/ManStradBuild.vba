'ToDo
' - [ ] Update Supplier Lookup
' - [ ] Code Documentation

Sub ManStradBuild()

'Prep
    If (ActiveSheet.AutoFilterMode And ActiveSheet.FilterMode) Or ActiveSheet.FilterMode Then ActiveSheet.ShowAllData 'Remove Filters
    Sheets("ManStrad").Cells.ClearContents 'Clear Cells

'Copy Data Over
    Sheets("ManStrad").Range("A:D").Value = Sheets("ManStructures").Range("A:D").Value 'Copy Data Over
    Columns("B:B").Delete Shift:=xlToLeft 'Delete Unneeded Column

'Headers
    Range("D1:F1").Value = Array("Component Requirement", "Supplier", "Comments") 'Fill Out Titles
    Range("G1").Formula = "=TODAY()-WEEKDAY(TODAY(),3)" 'Fill Out As Variable Date
    Range("H1:N1").FormulaR1C1 = "=RC[-1]+7" 'Continue Dates

    lrtarget = ActiveWorkbook.Sheets("ManStrad").Range("A1", Sheets("ManStrad").Range("A1").End(xlDown)).Rows.Count 'Find last row number

'Fill out various formulas
    Range("D2:D" & lrtarget).FormulaR1C1 = "=SUM(SUMIF('Reqs'!C[-2],RC[-3],'Reqs'!C),SUMIFS('Open Orders'!C[1],'Open Orders'!C[-2],RC[-3],'Open Orders'!C[-1],""Released""))"
    Range("E2:E" & lrtarget).FormulaR1C1 = "=VLOOKUP(RC[-3],'Suppliers'!C[-4]:C[-3],2,FALSE)"
    Range("G2:G" & lrtarget).FormulaR1C1 = "=SUM(SUMIFS('Reqs'!C4,'Reqs'!C2,RC1,'Reqs'!C3,""<=""&R1C+6),SUMIFS('Open Orders'!C5,'Open Orders'!C2,RC1,'Open Orders'!C4,""<=""&R1C+6,'Open Orders'!C3,""Released""))*RC3"
    Range("H2:N" & lrtarget).FormulaR1C1 = "=SUM(SUMIFS('Reqs'!C4,'Reqs'!C2,RC1,'Reqs'!C3,""<=""&R1C+6,'Reqs'!C3,"">=""&R1C),SUMIFS('Open Orders'!C5,'Open Orders'!C2,RC1,'Open Orders'!C4,""<=""&R1C+6,'Open Orders'!C3,""Released"",'Open Orders'!C4,"">=""&R1C))*RC3"

'Add filters
    Range("D1").AutoFilter
    ActiveWorkbook.Worksheets("ManStrad").AutoFilter.Sort.SortFields.Clear

'Sort Values Based on Component Requirement
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

    ActiveSheet.Range("$A1:$N" & lrtarget).AutoFilter Field:=4, Criteria1:="<1", Operator:=xlAnd 'Filter anything out that has a qty
    
    Application.DisplayAlerts = False 'Turn alerts off to avoid error flag
    Range("$A2:$N" & lrtarget).SpecialCells(xlCellTypeVisible).Delete 'Delete out visible cells
    Application.DisplayAlerts = True 'Turn alerts back on
    
'Cleanup
    ActiveSheet.Range("$A:$N").AutoFilter Field:=4 'Setup Filters
    Cells.EntireColumn.AutoFit 'Autofit Columns
    Range("A1").Select 'Reset Cursor

End Sub

