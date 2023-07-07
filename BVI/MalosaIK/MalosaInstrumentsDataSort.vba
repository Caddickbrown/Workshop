' ## To Do
' - [ ] Check if in correct exported sheet
' - [ ] Column Widths
' - [ ] Open Issues Log
' - [ ] Eventually Obselete with SQL
' - [ ] Variable Sheet Names in Requisitions Calculations



Sub MalosaInstrumentsDataSort()

Application.Calculation = xlManual
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

Dim Home As Workbook
Set Home = ThisWorkbook
Dim search As Range
Dim cnt As Integer
Dim colOrdr As Variant
Dim indx As Integer
Dim rng As Range

Dim LocationSheetName As String, RequisitionsSheetName As String, ReleasedOrdersSheetName As String, IPISSheetName As String, PivotSheetName As String, ShortageSheetName As String, IssueHighlightSheetName As String, PartNumberCalc As String, TotalRawMaterialQtyCalc As String, AMCOCalc As String, B1StockCalc As String, RMMaterialCalc As String, TotalReqForWeekCalc As String, RMShortageCalc As String, B1ShortageCalc As String, QuickReleaseCalc As String, ReleasedSOCalc As String, NetUsableRMCalc As String

'Tab Names
LocationSheetName = "Locations"
RequisitionsSheetName = "Requisition Demand"
ReleasedOrdersSheetName = "Released Shop Orders"
IPISSheetName = "IPIS"
PivotSheetName = "Requisitions Pivot"
ShortageSheetName = "Shortages"
IssueHighlightSheetName = "Issue Highlight"

'Formulas
'## Requisitions
WeekCalc = "=IF(D2<TODAY(),""Overdue"",CONCATENATE(YEAR(D2),"" - "",TEXT(ISOWEEKNUM(D2),""00"")))"
PCCalc = ""
IssueHighlightCalc = "=IF(COUNTIF('" & IssueHighlightSheetName & "'!A:A,B2)>0,""Issue"",""-"")"
RMCalc = "=IF(C2>VLOOKUP(B2,'" & LocationSheetName & "'!A:B,2,FALSE),""Insufficient RM"","""")"
SterilityCalc = "=IF(RIGHT(B2,1)=""S"",""Sterile"",""Non-Sterile"")"
RemainingCalc = "=COUNTA(A:A)-COUNTA(H:H)"
RemainingQtyCalc = "=SUM(C:C)-V6"
PCSpillCalc = "=IFERROR(SORT(UNIQUE(FILTER(F2:F999,F2:F999<>""""),FALSE,FALSE)),"""")"
SterileSpillCalc = "=SUMIFS($C:$C,$F:$F,IFERROR(SORT(UNIQUE(FILTER(F2:F999,F2:F999<>""""),FALSE,FALSE)),""""),$I:$I,N$2)"
NonSterileSpillCalc = "=SUMIFS($C:$C,$F:$F,IFERROR(SORT(UNIQUE(FILTER(F2:F999,F2:F999<>""""),FALSE,FALSE)),""""),$I:$I,O$2)"
TotalSpillCalc = "=IF(M3:M100="""","""",MMULT(IF(N3:O100="""",0,N3:O100),TRANSPOSE(SIGN(COLUMN(N2:O2)))))"

NoIssueReleasedCalc = "=SUMIFS($C:$C,$H:$H,$S3,$G:$G,""-"")"
IssueReleasedCalc ="=SUMIFS($C:$C,$H:$H,$S3,$G:$G,U$2)"
TotalReleasedCalc = "=SUM(T3:U3)"
PercentageReleasedCalc = "=V3/V6"

NoIssueReleaseCalc = "=SUMIFS($C:$C,$H:$H,$S4,$G:$G,""-"")"
IssueReleaseCalc ="=SUMIFS($C:$C,$H:$H,$S4,$G:$G,U$2)"
TotalReleaseCalc = "=SUM(T4:U4)"
PercentageReleaseCalc = "=V4/V6"

NoIssueNoReleaseCalc = "=SUMIFS($C:$C,$H:$H,$S5,$G:$G,""-"")"
IssueNoReleaseCalc ="=SUMIFS($C:$C,$H:$H,$S5,$G:$G,U$2)"
TotalNoReleaseCalc = "=SUM(T5:U5)"

TotalNoIssueCalc = "=SUM(T3:T5)"
TotalIssueCalc = "=SUM(U3:U5)"
TotalTotalCalc = "=SUM(V3:V5)"

WeekSpillCalc = "=IFERROR(SORT(UNIQUE(FILTER(E2:E999,E2:E999<>""""),FALSE,FALSE)),"""")"
TotalWeekSpillCalc = "=SUMIFS($C:$C,$E:$E,IFERROR(SORT(UNIQUE(FILTER(E2:E999,E2:E999<>""""),FALSE,FALSE)),""""))"

'## Locations
PartNumberCalc = "='" & PivotSheetName & "'!A2"
TotalRawMaterialQtyCalc = "=SUMIF(INDEX('" & IPISSheetName & "'!$A:$BZ,0,MATCH(""Part No"",'" & IPISSheetName & "'!$A$1:$BZ$1,0)),LEFT(A2,8)&""A"",INDEX('" & IPISSheetName & "'!$A:$BZ,0,MATCH(""On Hand Qty"",'" & IPISSheetName & "'!$A$1:$BZ$1,0)))"
AMCOCalc = "=SUMIFS(INDEX('" & IPISSheetName & "'!$A:$BZ,0,MATCH(""On Hand Qty"",'" & IPISSheetName & "'!$A$1:$BZ$1,0)),INDEX('" & IPISSheetName & "'!$A:$BZ,0,MATCH(""Warehouse"",'" & IPISSheetName & "'!$A$1:$BZ$1,0)),C$1,INDEX('" & IPISSheetName & "'!$A:$BZ,0,MATCH(""Part No"",'" & IPISSheetName & "'!$A$1:$BZ$1,0)),LEFT($A2,8)&""A"")"
B1StockCalc = "=E2+F2"
RMMaterialCalc = "=CONCATENATE(LEFT(A2,8),""A"")"
TotalReqForWeekCalc = "=SUMIFS('" & RequisitionsSheetName & "'!C:C,'" & RequisitionsSheetName & "'!B:B,A2)"
RMShortageCalc = "=B2-I2"
B1ShortageCalc = "=G2-I2"
QuickReleaseCalc = "=MIN(I2,G2-M2)"
ReleasedSOCalc = "=SUMIF('" & ReleasedOrdersSheetName & "'!A:A,A2,'" & ReleasedOrdersSheetName & "'!B:B)"
NetUsableRMCalc = "=G2-M2"

'Cut Data
    Sheets(1).Name = RequisitionsSheetName

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

'Add Tabs
    With Sheets
        .Add().Name = ShortageSheetName
        .Add().Name = LocationSheetName
        .Add().Name = ReleasedOrdersSheetName
        .Add().Name = IPISSheetName
        .Add().Name = IssueHighlightSheetName
    End With


'Fill Out Tabs
'Requisitions
    Sheets(RequisitionsSheetName).Select

    Range("E:XFD").ClearContents
    Range("E1:J1").Value = Array("Week", "PC", "Issue", "RM", "Sterility", "Notes")
    Range("E2:I2").Formula2 = Array(WeekCalc, PCCalc, IssueHighlightCalc, RMCalc, SterilityCalc)
    
    Range("E2:I" & Range("A" & Rows.Count).End(xlUp).Row).FillDown

    Range("D1").AutoFilter
    ActiveWorkbook.Worksheets(RequisitionsSheetName).AutoFilter.Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets(RequisitionsSheetName).AutoFilter.Sort. _
        SortFields.Add2 Key:=Range("D1:D" & Range("A" & Rows.Count).End(xlUp).Row), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(RequisitionsSheetName).AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Range("M1").Value = "Remaining"
    Range("N1:O1").Formula2 = Array(RemainingCalc,RemainingQtyCalc)

    Range("M2:P2").Value = Array("PC", "Sterile", "Non-Sterile", "Total")
    Range("M3:P3").Formula2 = Array(PCSpillCalc, SterileSpillCalc, NonSterileSpillCalc, TotalSpillCalc)

    Range("T2:V2").Value = Array("No Issue", "Issue", "Total")
    
    Range("S3").Value = "Released"
    Range("S4").Value = "To Release"
    Range("S5").Value = "Insufficient RM"
    Range("S6").Value = "Total"

    Range("T3:W3").Formula2 = Array(NoIssueReleasedCalc, IssueReleasedCalc, TotalReleasedCalc, PercentageReleasedCalc)
    Range("T4:W4").Formula2 = Array(NoIssueReleaseCalc, IssueReleaseCalc, TotalReleaseCalc, PercentageReleaseCalc)
    Range("T5:V5").Formula2 = Array(NoIssueNoReleaseCalc, IssueNoReleaseCalc, TotalNoReleaseCalc)
    Range("T6:V6").Formula2 = Array(TotalNoIssueCalc, TotalIssueCalc, TotalTotalCalc)

    Range("S10:T10").Value = Array("Week", "Total")
    Range("S11:T11").Formula2 = Array(WeekSpillCalc, TotalWeekSpillCalc)

'Pivot
    Range("C13").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Requisition Demand!R1C1:R999C10", Version:=8).CreatePivotTable _
        TableDestination:="", TableName:="PivotTable1", _
        DefaultVersion:=8
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
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Quantity"), "Sum of Quantity", xlSum
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Part No")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.Name = PivotSheetName

'Locations
    Sheets(LocationSheetName).Select

    Sheets(LocationSheetName).Rows("3:1048576").Delete Shift:=xlUp

    Sheets(LocationSheetName).Range("A1:N1").Value = Array("Part Number", "Total Raw Material Qty", "AMCO", "GOODS-IN", "INST&KNIVES", "CENTRAL-STORES", "B1 Stock", "RM Material", "Total Req For Week", "RM Shortage", "B1 Shortage", "Quick Release", "Released SO", "Net Usable RM")
    Sheets(LocationSheetName).Range("A2:N2").Formula2 = Array(PartNumberCalc, TotalRawMaterialQtyCalc, AMCOCalc, "", "", "", B1StockCalc, RMMaterialCalc, TotalReqForWeekCalc, RMShortageCalc, B1ShortageCalc, QuickReleaseCalc, ReleasedSOCalc, NetUsableRMCalc)
    Sheets(LocationSheetName).Range("C2").AutoFill Destination:=Range(Cells(2, 3), Cells(2, 6))

    Range("D1").AutoFilter
    Sheets(LocationSheetName).Range("A2").Select

'Shortages
'CONCATENATE(LEFT(ARRAY,8)&"A")


'Formatting
'Reqs
    Sheets(RequisitionsSheetName).Select
    Range("A1").Select
    With Columns("C:W")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "General"
    End With
    Columns("D:D").NumberFormat = "m/d/yyyy"
    Range("W3").NumberFormat = "0.00%"

    Cells.EntireColumn.AutoFit
    Range("V3:V6,T6:U6").Activate
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With

    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True

    Sheets(RequisitionsSheetName).Range("A1").Select
'Locations
    Sheets(LocationSheetName).Select
    Range("A1:N1").Font.Bold = True
    Set rng = Range("A1:N2")

    With rng.Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
        .LineStyle = xlContinuous
    
        .Item(xlDiagonalDown).LineStyle = xlNone
        .Item(xlDiagonalUp).LineStyle = xlNone
    End With

    With Range("A2,H1:I2").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    With Columns("A:N")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "General"
    End With

    Cells.EntireColumn.AutoFit

    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True

    Range("A2").Select

'Tabs
    With ActiveWorkbook.Sheets(ReleasedOrdersSheetName).Tab
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.599993896298105
    End With
    With ActiveWorkbook.Sheets(IPISSheetName).Tab
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.599993896298105
    End With
    With ActiveWorkbook.Sheets(IssueHighlightSheetName).Tab
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.599993896298105
    End With

'Sheet Order
    Sheets(RequisitionsSheetName).Move Before:=Sheets(1)
    Sheets(LocationSheetName).Move Before:=Sheets(2)
    Sheets(ShortageSheetName).Move Before:=Sheets(3)

    Sheets(RequisitionsSheetName).Select
    Range("A1").Select

Application.EnableEvents = True
Application.DisplayStatusBar = True
Application.ScreenUpdating = True
Application.Calculation = xlAutomatic

End Sub