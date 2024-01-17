Sub BVI_IK_Process()

    Range("C:C,E:E,F:F,G:J,L:L").Delete Shift:=xlToLeft
    Columns("F:F").Select
    Range(Selection, Selection.End(xlToRight)).Delete Shift:=xlToLeft
    Range("F1").FormulaR1C1 = "Format"
    Range("F2").FormulaR1C1 = _
        "=VLOOKUP(RC[-3],'[IK BVI Demand Plan.xlsm]SKUs'!C1:C2,2,FALSE)"
    Range("F2").AutoFill Destination:=Range("F2:F612")
    Range("J2").FormulaR1C1 = "=SUMIF(C[-8],RC[-1],C[-5])"
    Range("I2").FormulaR1C1 = "POOL"
    
    Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("C:C").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("E:E").Cut Destination:=Columns("C:C")
    Columns("E:E").Delete Shift:=xlToLeft
    Columns("F:H").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("J:J").ColumnWidth = 8.14
    Columns("J:J").Cut Destination:=Columns("G:G")
    Range("F1").FormulaR1C1 = "Brand"
    Range("H1").FormulaR1C1 = "Area"
    Range("A1").FormulaR1C1 = "Date"
    
    Range("A1").AutoFilter
    ActiveWorkbook.Worksheets(1).AutoFilter.Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets(1).AutoFilter.Sort. _
        SortFields.Add2 Key:=Range("G1:G612"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(1).AutoFilter. _
        Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Cells.EntireColumn.AutoFit
    
End Sub
