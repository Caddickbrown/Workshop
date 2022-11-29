Sub RefreshStats()

 'Copy yesterdays numbers up to the right field
    Sheets("Stats").Range("Q3:R3").Value = Sheets("Stats").Range("Q4:R4").Value 'Overwrite Yesterdays This Week figures
    Sheets("Stats").Range("Q6:R6").Value = Sheets("Stats").Range("Q7:R7").Value 'Overwrite Yesterdays Next Week figures

 'Change Cell to todays date
    Sheets("Stats").Range("P2").FormulaR1C1 = "=TODAY()" 'Update Date to be todays Date
    Sheets("Stats").Range("P2").Value = Sheets("Stats").Range("P2").Value 'Pastes todays date as values

 'Change Cell to todays date
    Sheets("HourStats").Range("P2").FormulaR1C1 = "=TODAY()" 'Update Date to be todays Date
    Sheets("HourStats").Range("P2").Value = Sheets("HourStats").Range("P2").Value 'Pastes todays date as values

 'Refresh PowerQuerys
    ActiveWorkbook.RefreshAll

 'Clean Up
    Sheets("Stats").Range("A1").Select 'Reset Cursor

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


Sub MondayFillTrackers()

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
 
 'Copy data to relevant tab (Next Weeks Tracker)
    lrtarget = ActiveWorkbook.Sheets("Order Well").Range("A1", Sheets("Order Well").Range("A1").End(xlDown)).Rows.Count 'Count rows on Order Well tab
    Sheets("Order Well").Range("A" & lrtarget + 1 & ":A" & lrtarget + 10).Value = Sheets("Stats").Range("P2").Value ' Measured Date
    Sheets("Order Well").Range("B" & lrtarget + 1 & ":B" & lrtarget + 10).Value = Sheets("Stats").Range("C2:C11").Value 'Weeks Out
    Sheets("Order Well").Range("C" & lrtarget + 1 & ":E" & lrtarget + 10).Value = Sheets("Stats").Range("D2:F11").Value 'BVI Qty
    Sheets("Order Well").Range("F" & lrtarget + 1 & ":H" & lrtarget + 10).Value = Sheets("Stats").Range("H2:J11").Value 'Malosa Qty
    Sheets("Order Well").Range("I" & lrtarget + 1 & ":K" & lrtarget + 10).Value = Sheets("HourStats").Range("D2:F11").Value 'BVI Hrs
    Sheets("Order Well").Range("L" & lrtarget + 1 & ":N" & lrtarget + 10).Value = Sheets("HourStats").Range("H2:J11").Value 'Malosa Hrs
 
 'Clean Up
    Range("A1").Select 'Reset Cursor

End Sub
