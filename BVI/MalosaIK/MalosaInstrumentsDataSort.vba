' ## To Do
' - [ ] Check if in correct exported sheet
' - [ ] Formatting for Spill "Table"
' - [ ] Column L Width Fix
' - [ ] Generate "Master Sheet"
' - [ ] Open Issues Log
' - [ ] Eventually Obselete with SQL
' 



Sub MalosaInstrumentsDataSort()

Dim Home As Workbook
Set Home = ThisWorkbook
Dim search As Range
Dim cnt As Integer
Dim colOrdr As Variant
Dim indx As Integer


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

    Range("E:XFD").ClearContents
    Range("E1:H1").Value = Array("Week", "PC", "Sterile", "Notes")
    Range("E2:G2").Formula2R1C1 = Array("=IF(RC[-1]<TODAY(),""Overdue"",CONCATENATE(YEAR(RC[-1]),"" - "",TEXT(ISOWEEKNUM(RC[-1]),""00"")))", "", "=IF(RIGHT(RC[-5],1)=""S"",""Sterile"",""Non-Sterile"")")
    
    Range("E2:G" & Range("A" & Rows.Count).End(xlUp).Row).FillDown

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

    Range("L2:O2").Value = Array("PC", "Sterile", "Non-Sterile", "Total")
    Range("L3:O3").Formula2 = Array("=IFERROR(SORT(UNIQUE(FILTER(F2:F999,F2:F999<>""""),FALSE,FALSE)),"""")", "=SUMIFS($C:$C,$F:$F,IFERROR(SORT(UNIQUE(FILTER(F2:F999,F2:F999<>""""),FALSE,FALSE)),""""),$G:$G,M$2)", "=SUMIFS($C:$C,$F:$F,IFERROR(SORT(UNIQUE(FILTER(F2:F999,F2:F999<>""""),FALSE,FALSE)),""""),$G:$G,N$2)", "=IF(L3:L100="""","""",MMULT(IF(M3:N100="""",0,M3:N100),TRANSPOSE(SIGN(COLUMN(M2:N2)))))")
    
    With Columns("L:O")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "General"
    End With
    
    Cells.EntireColumn.AutoFit
    
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Requisitions!R1C1:R339C4", Version:=8).CreatePivotTable _
        TableDestination:="Sheet1!R3C1", TableName:="PivotTable1", DefaultVersion _
        :=8
    Sheets("Sheet1").Select
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
    Sheets(1).Name = "Pivot"

    Sheets(2).Select
    Range("A1").Select

End Sub


