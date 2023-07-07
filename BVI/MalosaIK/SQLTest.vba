Sub SQLTest()
    Application.CommandBars("Queries and Connections").Visible = False
    ActiveWorkbook.Queries.Add Name:="InvPartInStockSQL", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Sql.Database(""bvi01sql02.database.windows.net"", ""BVI_Stage"", [Query=""SELECT *#(lf)FROM ifs.INVENTORY_PART_IN_STOCK_TAB#(lf)WHERE CONTRACT = '2051'#(lf)AND WAREHOUSE != 'Quality'#(lf)AND QTY_ONHAND > 0#(lf)AND AVAILABILITY_CONTROL_ID IS NOT NULL""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Source"
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=InvPartInStockSQL;Extended Properties=""""" _
        , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [InvPartInStockSQL]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "Table_InvPartInStockSQL"
        .Refresh BackgroundQuery:=False
    End With
End Sub


