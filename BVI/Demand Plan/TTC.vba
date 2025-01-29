Sub TTC()

'TTC
 
    Sheets("Main").Columns("A:A").TextToColumns Destination:=Range("Main[[#Headers],[SO Number]]"), _
        DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter _
        :=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, _
        Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True

    Sheets("Demand").Columns("A:A").TextToColumns Destination:=Range("Demand[[#Headers],[SO No]]"), _
        DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter _
        :=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, _
        Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True

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
   
    Sheets("Main").Select
 
'Archiving
 
    Calculate
    Application.Wait (Now() + TimeValue("00:00:10"))
 
    Dim DestinationTab As String
    Dim SourceTab As String
    Dim ArchiveDate As Date
   
    DestinationTab = "KPI"
    SourceTab = "Main"
    ArchiveDate = Format(Now(), "dd/mm/yyyy HH:nn")
 
    LastUsedRow = ActiveWorkbook.Sheets(DestinationTab).Range("A1", Sheets(DestinationTab).Range("A1").End(xlDown)).Rows.Count
 
    Sheets(DestinationTab).Range("A" & LastUsedRow + 1).Value = ArchiveDate 'Date/Time
    Sheets(DestinationTab).Range("B" & LastUsedRow + 1).Value = Sheets(SourceTab).Range("AG1").Value 'Blocked Lines
    Sheets(DestinationTab).Range("C" & LastUsedRow + 1).Value = Sheets(SourceTab).Range("AI1").Value 'Blocked Qty
    Sheets(DestinationTab).Range("D" & LastUsedRow + 1).Value = Sheets(SourceTab).Range("Z1").Value 'Lines to Check
    Sheets(DestinationTab).Range("E" & LastUsedRow + 1).Value = Sheets(SourceTab).Range("AA1").Value 'Qty to Check
    Sheets(DestinationTab).Range("F" & LastUsedRow + 1).Value = Sheets(SourceTab).Range("AK1").Value 'TW+ (This Week + x Weeks)
    Sheets(DestinationTab).Range("H" & LastUsedRow + 1).Value = Sheets(SourceTab).Range("AM1").Value 'Blocked Components
    Sheets(DestinationTab).Range("I" & LastUsedRow + 1).Value = Sheets(SourceTab).Range("AO1").Value 'Total Components
    Sheets(DestinationTab).Range("J" & LastUsedRow + 1).Value = Sheets(SourceTab).Range("AQ1").Value '% Material Available
    Sheets(DestinationTab).Range("K" & LastUsedRow + 1).Value = Sheets(SourceTab).Range("AS1").Value 'Hours Can't be Released
    Sheets(DestinationTab).Range("L" & LastUsedRow + 1).Value = Sheets(SourceTab).Range("AU1").Value 'Blocked Purchased SKUs
    Sheets(DestinationTab).Range("M" & LastUsedRow + 1).Value = Sheets(SourceTab).Range("AW1").Value 'Blocked Manufactured SKUs
    Sheets(DestinationTab).Range("N" & LastUsedRow + 1).Value = Sheets(SourceTab).Range("AY1").Value '% Unable
    
    Sheets(DestinationTab).Columns("A:A").NumberFormat = "dd/mm/yyyy"

    'ActiveWorkbook.Save
 
End Sub
 
'# Changelog

'## [2.0.0] - 2024-01-29

'### Added

'- New % Unable KPI

'### Removed

'- Demand Pivot References

'### Changed

'- Adjusted to the new Kit Planning Process

'## [1.2.2] - 2024-12-16
 
'### Added
 
'- Blocked Purchased SKUs
'- Blocked Manufactured SKUs
 
'## [1.2.1] - 2024-12-10
 
'### Added
 
'- Hours can't be released
 
'## [1.2.0] - 2024-12-09
 
'### Added
 
'- Blocked Components
'- Total Components
'% Material Available
 
'## [1.1.0] - 2024-10-10
 
'### Added
 
'- Lines to Check to Archive
'- Qty to Check to Archive
'- Changelog
 
'## [1.0.0] - 2024-10-08
 
'### Added
 
'- Initial Commit


