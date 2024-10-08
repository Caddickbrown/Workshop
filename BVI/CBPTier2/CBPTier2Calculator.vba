Sub CBPDailyPrep()

    Range("D3:D4").Value = Range("F3:F4").Value
    Range("E3:E4").ClearContents
    Range("J5:K6").ClearContents
    Range("A1").Select
    
End Sub

Sub CopyOver()

    Range("E3:E4").Value = Range("L5:L6").Value
    Range("A1").Select
    
End Sub

Sub ArchiveData()
    Dim DestinationTab As String
    Dim SourceTab As String
    Dim ArchiveDate As Date

    DestinationTab = "Archive"
    SourceTab = "Main"

    'Determine the archive date
    If Weekday(Date, vbMonday) = 1 Then
        ArchiveDate = Date - 3 'If today is Monday, use last Friday's date
    Else
        ArchiveDate = Date - 1 'Otherwise, use yesterday's date
    End If
   
    LastUsedRow = ActiveWorkbook.Sheets(DestinationTab).Range("A1", Sheets(DestinationTab).Range("A1").End(xlDown)).Rows.Count
   
    Sheets(DestinationTab).Range("A" & LastUsedRow + 1).Value = ArchiveDate 'Date
    Sheets(DestinationTab).Range("C" & LastUsedRow + 1).Value = Sheets(SourceTab).Range("F3").Value 'Kits
    Sheets(DestinationTab).Range("D" & LastUsedRow + 1).Value = Sheets(SourceTab).Range("F4").Value 'Instruments

    'ActiveWorkbook.Save

End Sub

'# Changelog

'## [1.1.2] - 2024-08-20

'### Added

'- Comments explaining Archive Info
'- ArchiveDate Calculation and Variable

'### Fixed

'- Changed Date Formula to use Fridays Date if on a Monday

'### Removed

'- Unused reference to "PastingRange"

'## [1.1.1] - 2024-08-16

'### Changed

'- Moved Order of Archive Columns

'## [1.1.0] - 2024-08-08

'### Added

'- Copy Over Macro

'## [1.0.0] - 2024-07-30

'### Added

'- Initial Commit
