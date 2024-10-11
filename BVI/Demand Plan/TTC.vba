Sub TTC()

'TTC

    Sheets("Demand").Columns("B:B").TextToColumns Destination:=Range("Demand[[#Headers],[Part No]]"), _
        DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter _
        :=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, _
        Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    
    Sheets("BOM Check").Columns("A:A").TextToColumns Destination:=Range("BOM_Check[[#Headers],[Part No]]") _
        , DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False _
        , Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
    Sheets("BOM Check").Columns("B:B").TextToColumns Destination:=Range( _
        "BOM_Check[[#Headers],[Component Part No]]"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
        
    Sheets("Hours").Columns("A:A").TextToColumns Destination:=Range( _
        "Hours[[#Headers],[PART_NO]]"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    
    Sheets("Demand Pivot").Select
    ActiveSheet.PivotTables("PivotTable2").PivotCache.Refresh
    Sheets("Main").Select

'Archiving

    Calculate
    Application.Wait (Now() + TimeValue("00:00:10"))

    Dim DestinationTab As String
    Dim SourceTab As String
    Dim ArchiveDate As Date
    
    DestinationTab = "KPI"
    SourceTab = "Main"
    ArchiveDate = TEXT(Now(), "dd/mm/yyyy hh:mm")

    LastUsedRow = ActiveWorkbook.Sheets(DestinationTab).Range("A1", Sheets(DestinationTab).Range("A1").End(xlDown)).Rows.Count

    Sheets(DestinationTab).Range("A" & LastUsedRow + 1).Value = ArchiveDate 'Date
    Sheets(DestinationTab).Range("B" & LastUsedRow + 1).Value = Sheets(SourceTab).Range("AG1").Value 'Blocked Lines
    Sheets(DestinationTab).Range("C" & LastUsedRow + 1).Value = Sheets(SourceTab).Range("AI1").Value 'Blocked Qty
    Sheets(DestinationTab).Range("D" & LastUsedRow + 1).Value = Sheets(SourceTab).Range("Z1").Value 'Lines to Check
    Sheets(DestinationTab).Range("E" & LastUsedRow + 1).Value = Sheets(SourceTab).Range("AA1").Value 'Qty to Check
    Sheets(DestinationTab).Range("F" & LastUsedRow + 1).Value = Sheets(SourceTab).Range("AK1").Value 'TW+ (This Week + x Weeks)

    'ActiveWorkbook.Save

End Sub

'# Changelog

'## [1.1.0] - 2024-10-10

'### Added

'- Lines to Check to Archive
'- Qty to Check to Archive
'- Changelog

'## [1.0.0] - 2024-10-08

'### Added

'- Initial Commit