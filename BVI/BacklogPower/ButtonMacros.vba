Sub RefreshStats()

 'Prep
    Sheets("Stats").Select 'Reset Sheet

 'Copy yesterdays numbers up to the right field
    Range("Q3:R3").Value = Range("Q4:R4").Value 'Overwrite Yesterdays This Week figures
    Range("Q6:R6").Value = Range("Q7:R7").Value 'Overwrite Yesterdays Next Week figures

 'Change Cell to todays date
    Range("P2").FormulaR1C1 = "=TODAY()" 'Update Date to be todays Date
    Range("P2").Value = Range("P2").Value 'Pastes todays date as values

 'Refresh PowerQuerys
    ActiveWorkbook.RefreshAll

 'Clean Up
    Range("A1").Select 'Reset Cursor

End Sub

Sub FillTrackers()

 'Prep
    Sheets("Stats").Select 'Reset Sheet

 'Copy data to relevant tab (This Weeks Tracker)
    lrtarget = ActiveWorkbook.Sheets("This Week Tracker").Range("A1", Sheets("This Week Tracker").Range("A1").End(xlDown)).Rows.Count 'Count rows on This Weeks tab
    Sheets("This Week Tracker").Range("A" & lrtarget + 1 & ":E" & lrtarget + 1).Value = Sheets("Stats").Range("M23:Q23").Value 'Pastes in This Weeks info 

 'Copy data to relevant tab (Daily Tracker)
    lrtarget = ActiveWorkbook.Sheets("Daily Tracker").Range("A1", Sheets("Daily Tracker").Range("A1").End(xlDown)).Rows.Count 'Count rows on Daily tab
    Sheets("Daily Tracker").Range("A" & lrtarget + 1 & ":E" & lrtarget + 1).Value = Sheets("Stats").Range("M26:Q26").Value 'Pastes in Daily info

 'Copy data to relevant tab (Next Weeks Tracker)
    lrtarget = ActiveWorkbook.Sheets("Next Week Tracker").Range("A1", Sheets("Next Week Tracker").Range("A1").End(xlDown)).Rows.Count 'Count rows on Next Week tab
    Sheets("Next Week Tracker").Range("A" & lrtarget + 1 & ":E" & lrtarget + 1).Value = Sheets("Stats").Range("M29:Q29").Value 'Pastes in Next Weeks info
 
 'Clean Up
    Range("A1").Select 'Reset Cursor

End Sub

Sub JustRefresh()

 'Prep
    Sheets("Stats").Select 'Reset Sheet

 'Refresh PowerQuerys
    ActiveWorkbook.RefreshAll

 'Clean Up
    Range("A1").Select 'Reset Cursor

End Sub
