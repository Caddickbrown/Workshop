Sub FillOutHistory()

   Dim LastUsedRow As Long

   ArchiveData DestinationTab:="Archive", SourceTabReleased:="SMW Dashboard", PastingRangeReleased:="N8:T8", SourceTabReqs:="Reqs Dashboard", PastingRangeReqs:="K10:P10"

   ResetView "SMW Dashboard"

   ActiveWorkbook.Save

End Sub

Sub ResetView(MainTab As String)
    Sheets(MainTab).Select
    Range("A1").Select
End Sub

Sub ArchiveData(DestinationTab As String, SourceTabReleased As String, PastingRangeReleased As String, SourceTabReqs As String, PastingRangeReqs As String)
   LastUsedRow = ActiveWorkbook.Sheets(DestinationTab).Range("A1", Sheets(DestinationTab).Range("A1").End(xlDown)).Rows.Count
   Sheets(DestinationTab).Range("A" & LastUsedRow + 1 & ":G" & LastUsedRow + 1).Value = Sheets(SourceTabReleased).Range(PastingRangeReleased).Value
   Sheets(DestinationTab).Range("H" & LastUsedRow + 1 & ":M" & LastUsedRow + 1).Value = Sheets(SourceTabReqs).Range(PastingRangeReqs).Value
End Sub

' # Changelog

' ## [1.0.1] - 2024-06-13

' ### Added

' - Save to End of FillOutHistory Macro
' - Changelog

' ## [1.0.0] - 

' ### Added

' - Initial Commit
' - Separate ResetView Macro
' - Separate Archive Macro

