Sub ArchiveBackOrder()

    Dim LastUsedRow As Long

    ArchiveData DestinationTab:="Log", SourceTab:="Pivots"
   
    ResetView "Pivots"

    ActiveWorkbook.Save

End Sub

Sub ResetView(MainTab As String)
    Sheets(MainTab).Select
    Range("A1").Select
End Sub

Sub ArchiveData(DestinationTab As String, SourceTab As String)

    Calculate

    LastUsedRow = ActiveWorkbook.Sheets(DestinationTab).Range("A1", Sheets(DestinationTab).Range("A1").End(xlDown)).Rows.Count
    Sheets(DestinationTab).Range("A" & LastUsedRow + 1 & ":E" & LastUsedRow + 1).Value = Sheets(SourceTab).Range("U4:Y4").Value

End Sub
