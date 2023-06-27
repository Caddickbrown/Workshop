'ToDo
'- [ ] Named Constants
'- [ ] Make Finding Last Row a Function
'- [ ] Looping?

Sub RefreshStats()

   AdjustPreviousValues "Stats" 
   AdjustPreviousValues "HourStats"

   UpdateDateToToday "Stats"
   UpdateDateToToday "HourStats"

   RefreshQueries
   ResetView

End Sub

Sub FillOutHistory()

   Dim LastUsedRow As Long
   
   CopyDataToHistoricalTab DestinationTab:="This Week Historical", PastingRange:="M23:Q23"
   CopyDataToHistoricalTab DestinationTab:="Daily Historical", PastingRange:="M26:Q26"
   CopyDataToHistoricalTab DestinationTab:="Next Week Historical", PastingRange:="M29:Q29"

   LastUsedRow = ActiveWorkbook.Sheets("Order Well").Range("A1", Sheets("Order Well").Range("A1").End(xlDown)).Rows.Count
  
   ArchiveData SourceDataTab:="Stats", PastingRangeStart:="A", PastingRangeEnd:="A", SourceDataLocation:="P2", LastUsedRow:=LastUsedRow         'Measured Date
   ArchiveData SourceDataTab:="Stats", PastingRangeStart:="B", PastingRangeEnd:="B", SourceDataLocation:="C2:C11", LastUsedRow:=LastUsedRow     'Weeks Out
   ArchiveData SourceDataTab:="Stats", PastingRangeStart:="C", PastingRangeEnd:="E", SourceDataLocation:="D2:F11", LastUsedRow:=LastUsedRow     'BVI Qty
   ArchiveData SourceDataTab:="Stats", PastingRangeStart:="F", PastingRangeEnd:="H", SourceDataLocation:="H2:J11", LastUsedRow:=LastUsedRow     'Malosa Qty
   ArchiveData SourceDataTab:="HourStats", PastingRangeStart:="I", PastingRangeEnd:="K", SourceDataLocation:="D2:F11", LastUsedRow:=LastUsedRow 'BVI Hrs
   ArchiveData SourceDataTab:="HourStats", PastingRangeStart:="L", PastingRangeEnd:="N", SourceDataLocation:="H2:J11", LastUsedRow:=LastUsedRow 'Malosa Hrs

   FillShipmentTracking DestinationTab:="Shipment Tracking", PastingRange:="M32:O32"

   ResetView

End Sub

Sub JustRefreshQueries()
   RefreshQueries
   ResetView
End Sub


Sub AdjustPreviousValues (ValuesTargetTab as String)
   Sheets(ValuesTargetTab).Range("Q3:R3").Value = Sheets(ValuesTargetTab).Range("Q4:R4").Value
   Sheets(ValuesTargetTab).Range("Q6:R6").Value = Sheets(ValuesTargetTab).Range("Q7:R7").Value
End Sub

Sub UpdateDateToToday (DatesTargetTab as String)
   Sheets(DatesTargetTab).Range("P2").FormulaR1C1 = "=TODAY()"
   Sheets(DatesTargetTab).Range("P2").Value = Sheets(DatesTargetTab).Range("P2").Value
End Sub

Sub CopyDataToHistoricalTab (DestinationTab as String, PastingRange as String)
   LastUsedRow = ActiveWorkbook.Sheets(DestinationTab).Range("A1", Sheets(DestinationTab).Range("A1").End(xlDown)).Rows.Count
   Sheets(DestinationTab).Range("A" & LastUsedRow + 1 & ":E" & LastUsedRow + 1).Value = Sheets("Stats").Range(PastingRange).Value
End Sub

Sub ArchiveData (SourceDataTab as String, PastingRangeStart as String, PastingRangeEnd as String, SourceDataLocation as String, ByRef LastUsedRow as Long)
   Sheets("Order Well").Range(PastingRangeStart & LastUsedRow + 1 & ":" & PastingRangeEnd & LastUsedRow + 10).Value = Sheets(SourceDataTab).Range(SourceDataLocation).Value
End Sub

Sub ResetView()
    Sheets("Stats").Select
    Range("A1").Select
End Sub

Sub RefreshQueries()
   ActiveWorkbook.RefreshAll
End Sub

Sub FillShipmentTracking (DestinationTab as String, PastingRange as String)
   LastUsedRow = ActiveWorkbook.Sheets(DestinationTab).Range("A1", Sheets(DestinationTab).Range("A1").End(xlDown)).Rows.Count
   Sheets(DestinationTab).Range("C" & LastUsedRow + 1 & ":E" & LastUsedRow + 1).Value = Sheets("Stats").Range(PastingRange).Value
End Sub
