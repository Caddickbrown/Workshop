Sub CBPDailyPrep()

    Range("D3:D5").Value = Range("F3:F5").Value
    Range("E3:E5").ClearContents
    Range("A1").Select
    
End Sub

Sub ArchiveData()
   Dim DestinationTab As String
   Dim SourceTab As String
   Dim ArchiveDate As Date
   
   DestinationTab = "Archive"
   SourceTab = "Main"
   
   ' Determine the archive date
   If Weekday(Date, vbMonday) = 1 Then
       ArchiveDate = Date - 3 ' If today is Monday, use last Friday's date
   Else
       ArchiveDate = Date - 1 ' Otherwise, use yesterday's date
   End If
   
   LastUsedRow = ActiveWorkbook.Sheets(DestinationTab).Range("A1", Sheets(DestinationTab).Range("A1").End(xlDown)).Rows.Count
   
   Sheets(DestinationTab).Range("A" & LastUsedRow + 1).Value = ArchiveDate
   Sheets(DestinationTab).Range("C" & LastUsedRow + 1).Value = Sheets(SourceTab).Range("F3").Value ' Area 1
   Sheets(DestinationTab).Range("D" & LastUsedRow + 1).Value = Sheets(SourceTab).Range("F4").Value ' Area 2
   Sheets(DestinationTab).Range("E" & LastUsedRow + 1).Value = Sheets(SourceTab).Range("F5").Value ' Cleanroom

   'ActiveWorkbook.Save

End Sub

' # Changelog

' ## [1.1.2] - 2024-08-20

' ### Added

' - Comments explaining Archive Info
' - ArchiveDate Calculation and Variable

' ### Fixed

' - Changed Date Formula to use Fridays Date if on a Monday

' ### Changed

' - Adjusted Macros for Vierpolders Template

' ### Removed

' - Unused reference to "PastingRange"
' - CopyOver Macro as not reqwuired for Vierpolders

' ## [1.1.1] - 2024-08-16

' ### Changed

' - Moved Order of Archive Columns

' ## [1.1.0] - 2024-08-08

' ### Added

' - Copy Over Macro

' ## [1.0.0] - 2024-07-30

' ### Added

' - Initial Commit


