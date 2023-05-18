Sub RefreshStats()

'Copy yesterdays numbers up to the right field
   YesterdayNumbers "Stats" 
   YesterdayNumbers "HourStats"

 'Change Cell to todays date
   ChangeDate "Stats"
   ChangeDate "HourStats"

 'Refresh PowerQuerys
    ActiveWorkbook.RefreshAll

 'Clean Up
    Sheets("Stats").Select 'Reset Sheet
    Range("A1").Select 'Reset Cursor

End Sub

Sub FillTrackers()

   TrackerFill shtTab3:= "This Week Tracker", tabLoc1:="M23:Q23"
   TrackerFill shtTab3:= "Daily Tracker", tabLoc1:="M26:Q26"
   TrackerFill shtTab3:= "Next Week Tracker", tabLoc1:="M29:Q29"

 'Clean Up
    Sheets("Stats").Select 'Reset Sheet
    Range("A1").Select 'Reset Cursor

End Sub

Sub JustRefresh()

 'Refresh PowerQuerys
    ActiveWorkbook.RefreshAll

 'Clean Up
    Sheets("Stats").Select 'Reset Sheet
    Range("A1").Select 'Reset Cursor

End Sub

Sub MondayFillTrackers()

   TrackerFill shtTab3:= "This Week Tracker", tabLoc1:="M23:Q23"
   TrackerFill shtTab3:= "Daily Tracker", tabLoc1:="M26:Q26"
   TrackerFill shtTab3:= "Next Week Tracker", tabLoc1:="M29:Q29"
 
 'Copy data to relevant tab (Next Weeks Tracker)
   lrtarget = ActiveWorkbook.Sheets("Order Well").Range("A1", Sheets("Order Well").Range("A1").End(xlDown)).Rows.Count 'Count rows on Order Well tab
   
   ArchiveData shtTab4:= "Stats", tabLoc2:= "A", tabLoc3:= "A", tabLoc4:= "P2" 'Measured Date
   ArchiveData shtTab4:= "Stats", tabLoc2:= "B", tabLoc3:= "B", tabLoc4:= "C2:C11" 'Weeks Out
   ArchiveData shtTab4:= "Stats", tabLoc2:= "C", tabLoc3:= "E", tabLoc4:= "D2:F11" 'BVI Qty
   ArchiveData shtTab4:= "Stats", tabLoc2:= "F", tabLoc3:= "H", tabLoc4:= "H2:J11" 'Malosa Qty
   ArchiveData shtTab4:= "HourStats", tabLoc2:= "I", tabLoc3:= "K", tabLoc4:= "D2:F11" 'BVI Hrs
   ArchiveData shtTab4:= "HourStats", tabLoc2:= "L", tabLoc3:= "N", tabLoc4:= "H2:J11" 'Malosa Hrs
 
 'Clean Up
    Sheets("Stats").Select 'Reset Sheet
    Range("A1").Select 'Reset Cursor

End Sub

Sub YesterdayNumbers (shtTab1 as String)
'Copy yesterdays numbers or hours up to the right field
   Sheets(shtTab1).Range("Q3:R3").Value = Sheets(shtTab1).Range("Q4:R4").Value 'Overwrite Yesterdays "This Week" figures
   Sheets(shtTab1).Range("Q6:R6").Value = Sheets(shtTab1).Range("Q7:R7").Value 'Overwrite Yesterdays "Next Week" figures
End Sub

Sub ChangeDate (shtTab2 as String)
'Change Cell to todays date
   Sheets(shtTab2).Range("P2").FormulaR1C1 = "=TODAY()" 'Update Date to be todays Date
   Sheets(shtTab2).Range("P2").Value = Sheets(shtTab2).Range("P2").Value 'Pastes todays date as values
End Sub

Sub TrackerFill (shtTab3 as String, tabLoc1 as String)
   lrtarget = ActiveWorkbook.Sheets(shtTab3).Range("A1", Sheets(shtTab3).Range("A1").End(xlDown)).Rows.Count 'Count rows on This Weeks tab
   Sheets(shtTab3).Range("A" & lrtarget + 1 & ":E" & lrtarget + 1).Value = Sheets("Stats").Range(tabLoc1).Value 'Pastes in This Weeks info 
End Sub

Sub ArchiveData (shtTab4 as String, tabLoc2 as String, tabLoc3 as String, tabLoc4 as String)
   Sheets("Order Well").Range(tabLoc2 & lrtarget + 1 & ":" & tabLoc3 & lrtarget + 10).Value = Sheets(shtTab4).Range(tabLoc4).Value ' Measured Date
End Sub