Sub Clear_Sheet()
'
' Clear_Sheet Macro
'

'
    Range("E3:I1048576").ClearContents
    Range("A1").Select
End Sub
Sub Process()
'
' Process Macro
'

'
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    Range("E2:I2").AutoFill Destination:=Range("E2:I" & last_row)

    Columns("A:I").Select
    Selection.Copy
    Sheets("Output").Select
    Columns("A:A").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B1").Select
    Application.CutCopyMode = False
    
    If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
    End If
    
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("Output").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Output").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("B1:B93"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Output").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    
    Range("A1").Select

End Sub
