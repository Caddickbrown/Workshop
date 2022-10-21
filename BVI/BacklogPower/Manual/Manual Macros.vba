Sub Prep_Sheet()

'Prep
    Application.ScreenUpdating = False 'Cleans View up a bit, so it doesn't jump around
    Sheets("Stats").Select 'Reset Sheet

'Copy yesterdays numbers up to the right field
    Range("Q3:R3").Value = Range("Q4:R4").Value 'Overwrite Yesterdays This Week figures
    Range("Q6:R6").Value = Range("Q7:R7").Value 'Overwrite Yesterdays Next Week figures

'Change Cell to todays date
    Range("P2").FormulaR1C1 = "=TODAY()" 'Update Date to be todays Date
    Range("P2").Value = Range("P2").Value

'Clean Up
    Range("A1").Select 'Reset Cursor
    Application.ScreenUpdating = True 'Reset Screen Updating

End Sub

Sub Archive()

   Application.ScreenUpdating = False ' Cleans View up a bit, so it doesn't jump around

' Make sure you're on the right sheet
    Sheets("Stats").Select ' Reset Sheet

' Copy data to relevant tab (This Week)
    Range("M23:Q23").Copy ' Copy This Weeks data
    Sheets("Archive").Select ' Move to "Archive" Tab
    lrtarget = Cells.Find("*", Cells(1, 1), xlFormulas, xlPart, xlByRows, xlPrevious, False).Row ' Find the last row on the sheet
    Cells(lrtarget + 1, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False ' Paste as Values
    Sheets("Stats").Select ' Reset Sheet

' Copy data to relevant tab (Daily)
    Range("N26:Q26").Copy ' Copy Daily data
    Sheets("Archive").Select ' Move to "Archive" Tab
    lrtarget = Cells.Find("*", Cells(1, 1), xlFormulas, xlPart, xlByRows, xlPrevious, False).Row ' Find the last row on the sheet
    Cells(lrtarget, 6).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False ' Paste as Values
    Sheets("Stats").Select ' Reset Sheet

' Copy data to relevant tab (Next Week)
    Range("N29:Q29").Copy ' Copy Next Weeks data
    Sheets("Archive").Select ' Move to "Archive" Tab
    lrtarget = Cells.Find("*", Cells(1, 1), xlFormulas, xlPart, xlByRows, xlPrevious, False).Row ' Find the last row on the sheet
    Cells(lrtarget, 10).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False ' Paste as Values
    Range("A1").Select ' Reset Cursor
    Sheets("Stats").Select ' Reset Sheet
    
    Range("A1").Select ' Reset Cursor
    Application.ScreenUpdating = True ' Reset Screen Updating

End Sub