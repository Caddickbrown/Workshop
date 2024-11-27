Sub BVI_IK_Process_v4()

' Define variables
Dim ws As Worksheet
Dim search As Range
Dim cnt As Integer
Dim colOrdr As Variant
Dim indx As Integer

ReleasedOrdersSheetName = "Released Shop Orders"

' Display a confirmation message
    Dim response As VbMsgBoxResult
    response = MsgBox("Are you sure you want to run this macro?", vbYesNo + vbQuestion, "Confirmation")

    ' Check the user's response
    If response = vbNo Then
        MsgBox "Macro cancelled.", vbInformation, "Cancelled"
        Exit Sub
    End If

' Filter to important data

    ' Sort column order
    Sheets(1).Name = ReleasedOrdersSheetName

    Set ws = ActiveSheet

    colOrdr = Array("Order No", "Part No", "Priority Category", "Start Date", "Lot Size") 'define column order with header names here

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


    ' Delete the rows after the last row in the dataset
    If cnt - 1 < ws.Columns.Count Then
        ws.Range(ws.Columns(cnt), ws.Columns(ws.Columns.Count)).Delete
    End If

' Add additional Columns

    Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("F:H").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

' Enter Values
    ' Columns
    Range("A1").Value = "Date"
    Range("F1:H1").Value = Array("Brand", "Format", "Area")
    Range("J1").Value = "Hours"

    ' Summary Section
    Range("N1:O1").Value = Array("Qty", "Hrs")
    Range("M2").Value = "POOL"

    ' Formulas
    Range("D2").Formula = "=CONCATENATE(TEXT(ISOWEEKNUM(E2),""00""),TEXT(WEEKDAY(E2,2),""00""))" ' Brand
    Range("F2").Formula = "=SWITCH(LEFT(C2,4),""MMSU"",""Malosa"",""BVI"")" ' Brand
    Range("G2").Formula = "=SWITCH(LEFT(C2,4),""MMSU"",IF(RIGHT(C2,1)=""S"",""Shelf"",""Kit""),IFERROR(VLOOKUP($C2,'[IK BVI Demand Plan.xlsm]SKUs'!$A:$B,2,FALSE),NA()))" ' Format
    ' Need to add Brand and Area Formulas
    Range("J2").Formula = "=SUMIFS('[Instruments Daily Plan.xlsm]Hrs'!$C:$C,'[Instruments Daily Plan.xlsm]Hrs'!$A:$A,$C2,'[Instruments Daily Plan.xlsm]Hrs'!D:D,""<>""&""Boxing"")*$I2" ' Hours
    Range("N2").Formula = "=SUMIF($D:$D,$M2,I:I)"
    Range("O2").Formula = "=SUMIF($D:$D,$M2,J:J)"
    Range("O6").Formula = "=IF(N6="""",""<- Paste Concat Here"",CONCATENATE(N6,"";"",O5))" ' Concatenate

    LastUsedRow = ActiveSheet.Range("B1", ActiveSheet.Range("B1").End(xlDown)).Rows.Count

    Range("D2").AutoFill Destination:=Range("D2:D" & LastUsedRow)
    Range("F2:G2").AutoFill Destination:=Range("F2:G" & LastUsedRow)
    Range("J2").AutoFill Destination:=Range("J2:J" & LastUsedRow)
    
    ' Filters
    Range("A1").AutoFilter
    ActiveWorkbook.Worksheets(1).AutoFilter.Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets(1).AutoFilter.Sort. _
        SortFields.Add2 Key:=Range("E1:E612"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(1).AutoFilter. _
        Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

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

    ActiveWorkbook.Worksheets(1).AutoFilter.Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets(1).AutoFilter.Sort. _
        SortFields.Add2 Key:=Range("F1:F612"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(1).AutoFilter. _
        Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ' Final Formatting
    With Columns("A:O")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("J2").Select
    Range(Selection, Selection.End(xlDown)).NumberFormat = "0.00"
    Range("F:G").NumberFormat = "General"
    Range("O2").NumberFormat = "0.00"
    Cells.EntireColumn.AutoFit
    Columns("N:N").ColumnWidth = 8.14

End Sub

' # Changelog

' ## [4.0.0] - 2024-11-27

' ### Added

' - Confirmation Step at Start

' ## [3.0.0] - 2024-10-22

' ### Added

' - Formula to calculate Priority Category
' - Autofill Priority Category

' ### Fixed

' - Missing Close Bracket in Formula

' ## [1.0.1] - 2024-07-04

' ### Added

' - Changelog

' ### Fixed

' - Removed Boxing hours from Calculation as causing double count.