' ## To Do
' - [ ] Check if in correct exported sheet
' - [ ] Column Widths
' - [ ] Generate "Master Sheet"
' - [ ] Open Issues Log
' - [ ] Extend Pivot Table
' - [ ] Eventually Obselete with SQL
' - [ ] Tidy Up code



Sub MalosaInstrumentsDataSort()

Dim Home As Workbook
Set Home = ThisWorkbook
Dim search As Range
Dim cnt As Integer
Dim colOrdr As Variant
Dim indx As Integer


Dim MainSheetName As String, RequisitionsSheetName As String, ReleasedOrdersSheetName As String, IPISSheetName As String, PartNumberCalc As String, TotalRawMaterialQtyCalc As String, AMCOCalc As String, B1StockCalc As String, RMMaterialCalc As String, TotalReqForWeekCalc As String, RMShortageCalc As String, B1ShortageCalc As String, QuickReleaseCalc As String, ReleasedSOCalc As String, NetUsableRMCalc As String

'Tab Names
MainSheetName = "Locations"
RequisitionsSheetName = "Requisition Demand"
ReleasedOrdersSheetName = "Released Shop Orders"
IPISSheetName = "IPIS"

'Formulas
PartNumberCalc = "='" & RequisitionsSheetName & "'!A2"
TotalRawMaterialQtyCalc = "=SUMIF(INDEX(" & IPISSheetName & "!$A:$BZ,0,MATCH(""Part No""," & IPISSheetName & "!$A$1:$BZ$1,0)),LEFT(A2,8)&""A"",INDEX(" & IPISSheetName & "!$A:$BZ,0,MATCH(""On Hand Qty""," & IPISSheetName & "!$A$1:$BZ$1,0)))"
AMCOCalc = "=SUMIFS(INDEX(" & IPISSheetName & "!$A:$BZ,0,MATCH(""On Hand Qty""," & IPISSheetName & "!$A$1:$BZ$1,0)),INDEX(" & IPISSheetName & "!$A:$BZ,0,MATCH(""Warehouse""," & IPISSheetName & "!$A$1:$BZ$1,0)),C$1,INDEX(" & IPISSheetName & "!$A:$BZ,0,MATCH(""Part No""," & IPISSheetName & "!$A$1:$BZ$1,0)),LEFT($A2,8)&""A"")"
B1StockCalc = "=E2+F2"
RMMaterialCalc = "=CONCATENATE(LEFT(A2,8),""A"")"
TotalReqForWeekCalc = "=VLOOKUP(A2,'" & RequisitionsSheetName & "'!A:B,2,0)"
RMShortageCalc = "=B2-I2"
B1ShortageCalc = "=G2-I2"
QuickReleaseCalc = "=MIN(I2,G2-M2)"
ReleasedSOCalc = "=SUMIF('" & ReleasedOrdersSheetName & "'!A:A,A2,'" & ReleasedOrdersSheetName & "'!B:B)"
NetUsableRMCalc = "=G2-M2"




    Sheets(1).Name = "Requisitions"

    colOrdr = Array("Requisition ID", "Part No", "Quantity", "Proposed Start Date") 'define column order with header names here

    cnt = 1

    For indx = LBound(colOrdr) To UBound(colOrdr)
        Set search = Rows("1:1").Find(colOrdr(indx), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
        If Not search Is Nothing Then
            If search.Column <> cnt Then
                search.EntireColumn.Cut
                Columns(cnt).Insert Shift:=xlToRight
                Application.CutCopyMode = False
            End If
        cnt = cnt + 1
        End If
    Next indx

    Sheets.Add.Name = "MPKG" 'MPKG
    Sheets(1).Select

    Range("E:XFD").ClearContents
    Range("E1:J1").Value = Array("Week", "PC", "MPKG", "RM", "Sterility", "Notes")
    Range("E2:I2").Formula2 = Array("=IF(D2<TODAY(),""Overdue"",CONCATENATE(YEAR(D2),"" - "",TEXT(ISOWEEKNUM(D2),""00"")))", "", "=IF(COUNTIF(MPKG!A:A,B2)>0,""Issue"",""-"")", "", "=IF(RIGHT(B2,1)=""S"",""Sterile"",""Non-Sterile"")")
    
    Range("E2:I" & Range("A" & Rows.Count).End(xlUp).Row).FillDown

    Range("D1").AutoFilter
    ActiveWorkbook.Worksheets("Requisitions").AutoFilter.Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Requisitions").AutoFilter.Sort. _
        SortFields.Add2 Key:=Range("D1:D" & Range("A" & Rows.Count).End(xlUp).Row), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Requisitions").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Range("M1").Value = "Remaining"
    Range("N1").Formula2 = "=COUNTA(A:A)-COUNTA(H:H)"

    Range("M2:P2").Value = Array("PC", "Sterile", "Non-Sterile", "Total")
    Range("M3:P3").Formula2 = Array("=IFERROR(SORT(UNIQUE(FILTER(F2:F999,F2:F999<>""""),FALSE,FALSE)),"""")", "=SUMIFS($C:$C,$F:$F,IFERROR(SORT(UNIQUE(FILTER(F2:F999,F2:F999<>""""),FALSE,FALSE)),""""),$I:$I,N$2)", "=SUMIFS($C:$C,$F:$F,IFERROR(SORT(UNIQUE(FILTER(F2:F999,F2:F999<>""""),FALSE,FALSE)),""""),$I:$I,O$2)", "=IF(M3:M100="""","""",MMULT(IF(N3:O100="""",0,N3:O100),TRANSPOSE(SIGN(COLUMN(N2:O2)))))")
    
    With Columns("C:W")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "General"
    End With
    Columns("D:D").NumberFormat = "m/d/yyyy"
    Range("W3").NumberFormat = "0.00%"
    
    Range("T2:V2").Value = Array("No Issue", "Issue", "Total")
    Range("S3").Value = Array("To Release")
    Range("S4").Value = Array("Insufficient RM")
    Range("S5").Value = Array("Total")

    Range("T3:W3").Formula2 = Array("=SUMIFS($C:$C,$H:$H,$S3,$G:$G,""-"")", "=SUMIFS($C:$C,$H:$H,$S3,$G:$G,U$2)", "=SUM(T3:U3)", "=V3/V5")
    Range("T4:W4").Formula2 = Array("=SUMIFS($C:$C,$H:$H,$S4,$G:$G,""-"")", "=SUMIFS($C:$C,$H:$H,$S4,$G:$G,U$2)", "=SUM(T4:U4)", "")
    Range("T5:W5").Formula2 = Array("=SUM(T3:T4)", "=SUM(U3:U4)", "=SUM(V3:V4)", "")

    Range("S10:T10").Value = Array("Week", "Total")
    Range("S11:T11").Formula2 = Array("=IFERROR(SORT(UNIQUE(FILTER(E2:E942,E2:E942<>""""),FALSE,FALSE)),"""")" ,"=SUMIFS($C:$C,$E:$E,IFERROR(SORT(UNIQUE(FILTER(E2:E942,E2:E942<>""""),FALSE,FALSE)),""""))")

    Cells.EntireColumn.AutoFit
    Range("V3:V5,T5:U5").Activate
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With

    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Requisitions!R1C1:R999C10", Version:=8).CreatePivotTable _
        TableDestination:="Sheet2!R1C1", TableName:="PivotTable1", DefaultVersion _
        :=8
    Sheets("Sheet2").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable1")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable1").RepeatAllLabels xlRepeatLabels
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Part No").Orientation = xlRowField
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables("PivotTable1").PivotFields("Quantity"), "Sum of Quantity", xlSum
    Sheets(1).Name = "Requisition Demand"

    Sheets(2).Select
    Range("A1").Select

    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True




    Sheets.Add.Name = "Shortages" 'StockChecker(Locations)
    Sheets.Add.Name = MainSheetName 'StockChecker(Locations)
    Sheets.Add.Name = ReleasedOrdersSheetName 'Released Orders
    Sheets.Add.Name = IPISSheetName 'IPIS


'Shortages
'CONCATENATE(LEFT(ARRAY,8)&"A")

'Locations
Sheets(MainSheetName).Select

Sheets(MainSheetName).Rows("3:1048576").Delete Shift:=xlUp
Sheets(RequisitionsSheetName).Rows("2:1048576").ClearContents
Sheets(ReleasedOrdersSheetName).Rows("2:1048576").ClearContents
Sheets(IPISSheetName).Cells.ClearContents

Sheets(MainSheetName).Range("A1:N1").Value = Array("Part Number", "Total Raw Material Qty", "AMCO", "GOODS-IN", "INST&KNIVES", "CENTRAL-STORES", "B1 Stock", "RM Material", "Total Req For Week", "RM Shortage", "B1 Shortage", "Quick Release", "Released SO", "Net Usable RM")
Sheets(MainSheetName).Range("A2:N2").Formula2 = Array(PartNumberCalc,TotalRawMaterialQtyCalc, AMCOCalc, "", "", "", B1StockCalc, RMMaterialCalc, TotalReqForWeekCalc, RMShortageCalc, B1ShortageCalc, QuickReleaseCalc, ReleasedSOCalc, NetUsableRMCalc)
Sheets(MainSheetName).Range("C2").AutoFill Destination:=Range(Cells(2, 3), Cells(2, 6))

Sheets(RequisitionsSheetName).Range("A1:C1").Value = Array("Part Numbers", "Sum of Quantity", "Priority")

Sheets(ReleasedOrdersSheetName).Range("A1:B1").Value = Array("Part Numbers", "Lot Size", "Priority")

Sheets(MainSheetName).Range("A2").Select

'Released Shop Orders
'Data Dump

'IPIS
'Data Dump


'MPKG
'Data Dump
'Include Formula to look up to it


End Sub
