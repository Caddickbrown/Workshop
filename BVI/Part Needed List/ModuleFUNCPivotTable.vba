'WS is data source, WB is home workbook, out is where to place the pvot table, PT is the pivot table
Function CreatePIvot(WS As Worksheet, WB As Workbook, out As Worksheet, PT As PivotTable)

    Dim RowCount As Long, ColumnCount As Long, IFSdata As Range
        
    'clears any pivot table currently on sheet as you can't place a new one on top of it
    out.Cells.Clear
    
    'finds teh edges of the data
    RowCount = WS.UsedRange.Rows.Count
    ColumnCount = WS.Cells(2, Columns.Count).End(xlToLeft).Column
    
    'sets the range to turn into a pivot
    Set IFSdata = Range(WS.Cells(1, 1), WS.Cells(RowCount, ColumnCount))
    
    'Creates the pivot table with teh specified data
    Set PT = WB.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=IFSdata).CreatePivotTable(TableDestination:=out.Cells(1, 1))
    
    'adds the required data to the pivot, part number then revision 1st then count and qty in teh values
    
End Function