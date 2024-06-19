Sub RefreshDetail()

Dim wb As Workbook

' Open Plans
Application.Workbooks.Open ("https://bvx.sharepoint.com/Operations/Reporting%20and%20KPIs/Schedules/Daily%20Plan.xlsm")
Application.Workbooks.Open ("https://bvx.sharepoint.com/Operations/Reporting%20and%20KPIs/Schedules/Instruments%20Daily%20Plan.xlsm")

' Wait to load and calculate
Application.Wait (Now() + TimeValue("00:00:15"))

'Activate and Calculate
Workbooks("Daily Plan.xlsm").Activate
Calculate
Workbooks("Instruments Daily Plan.xlsm").Activate
Calculate
Workbooks("Backorder Trending.xlsm").Activate
Calculate

' Archive
ArchiveData DestinationTab:="Archive", SourceTab:="Data", PastingRange:="A2:K2"

' Close Workbooks
Workbooks("Daily Plan.xlsm").Close SaveChanges:=False
Workbooks("Instruments Daily Plan.xlsm").Close SaveChanges:=False

' Save
ActiveWorkbook.Save

End Sub

Sub ArchiveData(DestinationTab As String, SourceTab As String, PastingRange As String)
   LastUsedRow = ActiveWorkbook.Sheets(DestinationTab).Range("A1", Sheets(DestinationTab).Range("A1").End(xlDown)).Rows.Count
   Sheets(DestinationTab).Range("A" & LastUsedRow + 1 & ":K" & LastUsedRow + 1).Value = Sheets(SourceTab).Range(PastingRange).Value
End Sub

' # Changelog

' ## [1.1.0] - 2024-06-19

' ### Added

' - Ensure Calculations are Complete for Plan Workbooks

' ## [1.0.1] - 2024-06-13

' ### Changed

' - Consolidated Date and Time into single field for simplicity

' ### Added

' - Save to End of RefreshDetail Macro

' ## [1.0.0] - 2024-06-13

' ### Added

' - Initial Commit
' - Separate Archive Macro