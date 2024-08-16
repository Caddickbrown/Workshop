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
   Dim PastingRange As String

   DestinationTab = "Archive"
   SourceTab = "Main"
   PastingRange = "F3:F4"

   LastUsedRow = ActiveWorkbook.Sheets(DestinationTab).Range("A1", Sheets(DestinationTab).Range("A1").End(xlDown)).Rows.Count
   Sheets(DestinationTab).Range("A" & LastUsedRow + 1).Value = Date - 1
   Sheets(DestinationTab).Range("C" & LastUsedRow + 1).Value = Sheets(SourceTab).Range("F3").Value
   Sheets(DestinationTab).Range("D" & LastUsedRow + 1).Value = Sheets(SourceTab).Range("F4").Value

   'ActiveWorkbook.Save

End Sub

' # Changelog

' ## [1.1.1] - 2024-08-16

' ### Changed

' - Moved Order of Archive Columns

' ## [1.1.0] - 2024-08-08

' ### Added

' - Copy Over Macro

' ## [1.0.0] - 2024-07-30

' ### Added

' - Initial Commit


