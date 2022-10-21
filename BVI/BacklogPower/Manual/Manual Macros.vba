Sub Prep_Sheet()

'Prep
    Application.ScreenUpdating = False 'Cleans View up a bit, so it doesn't jump around
    Sheets("Stats").Select 'Reset Sheet

'Copy yesterdays numbers up to the right field
    Range("Q3:R3").Value = Range("Q4:R4").Value 'Overwrite Yesterdays This Week figures
    Range("Q6:R6").Value = Range("Q7:R7").Value 'Overwrite Yesterdays Next Week figures

'Change Cell to todays date
    Range("P2").FormulaR1C1 = "=TODAY()" 'Update Date to be todays Date
    Range("P2").Value = Range("P2").Value 'Paste date as values

'Clean Up
    Range("A1").Select 'Reset Cursor
    Application.ScreenUpdating = True 'Reset Screen Updating

End Sub

Sub Archive()

' Make sure you're on the right sheet
    Sheets("Stats").Select ' Reset Sheet

' Copy data to Archive tab
    lrtarget = ActiveWorkbook.Sheets("Archive").Range("A1", Sheets("Archive").Range("A1").End(xlDown)).Rows.Count 'Count rows on Archive tab 
    Sheets("Archive").Range("A" & lrtarget + 1 & ":E" & lrtarget + 1).Value = Sheets("Stats").Range("M23:Q23").Value 'Pastes in This Week Info
    Sheets("Archive").Range("F" & lrtarget + 1 & ":I" & lrtarget + 1).Value = Sheets("Stats").Range("N26:Q26").Value 'Pastes in Daily Info
    Sheets("Archive").Range("J" & lrtarget + 1 & ":M" & lrtarget + 1).Value = Sheets("Stats").Range("N29:Q29").Value 'Pastes in Next Week Info
    
    Range("A1").Select ' Reset Cursor

End Sub