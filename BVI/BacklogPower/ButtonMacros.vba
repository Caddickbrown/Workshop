Sub RefreshStats()

   Application.ScreenUpdating = False ' Cleans View up a bit, so it doesn't jump around

' Make sure you're on the right sheet
    Sheets("Stats").Select ' Reset Sheet

' Copy yesterdays numbers up to the right field
    Range("Q4").Copy ' Copy BVI This Week Figure (Which is now yesterdays)
    Range("Q3").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False ' Paste as Values in Yesterdays Row
    Range("R4").Copy ' Copy Malosa This Week Figure (Which is now yesterdays)
    Range("R3").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False ' Paste as Values in Yesterdays Row
        
' Copy yesterdays numbers up to the right field
    Range("Q7").Copy ' Copy BVI Next Week Figure (Which is now yesterdays)
    Range("Q6").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False ' Paste as Values in Yesterdays Row
    Range("R7").Copy ' Copy Malosa Next Week Figure (Which is now yesterdays)
    Range("R6").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False ' Paste as Values in Yesterdays Row

' Change to todays date
    Range("P2").FormulaR1C1 = "=TODAY()" ' Update Date to be todays Date
    Range("P2").Select ' Select cell (Need to for the paste)
    Selection.Copy ' Copy cell
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False ' Paste as Values - stopping it from automatically updating tomorrow
    
' Refresh PowerQuerys
    ActiveWorkbook.RefreshAll

    Range("A1").Select ' Reset Cursor
    Application.ScreenUpdating = True ' Reset Screen Updating

End Sub

Sub FillTrackers()

   Application.ScreenUpdating = False ' Cleans View up a bit, so it doesn't jump around

' Make sure you're on the right sheet
    Sheets("Stats").Select ' Reset Sheet

' Copy data to relevant tab (This Weeks Tracker)
    Range("M23:Q23").Copy ' Copy Todays "This Week" Info
    Sheets("This Week Tracker").Select ' Move to "This Week Tracker" sheet
    lrtarget = Cells.Find("*", Cells(1, 1), xlFormulas, xlPart, xlByRows, xlPrevious, False).Row ' Find the last row on the sheet
    Cells(lrtarget + 1, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False ' Paste as Values on next row down
    Range("A1").Select ' Reset Cursor
    Sheets("Stats").Select ' Reset Sheet

' Copy data to relevant tab (Daily Tracker)
    Range("M26:Q26").Copy ' Copy Todays "Daily" Info
    Sheets("Daily Tracker").Select ' Move to "Daily Tracker" sheet
    lrtarget = Cells.Find("*", Cells(1, 1), xlFormulas, xlPart, xlByRows, xlPrevious, False).Row ' Find the last row on the sheet
    Cells(lrtarget + 1, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False ' Paste as Values on next row down
    Range("A1").Select ' Reset Cursor
    Sheets("Stats").Select ' Reset Sheet

' Copy data to relevant tab (Next Weeks Tracker)
    Range("M29:Q29").Copy ' Copy Todays "Next Week" Info
    Sheets("Next Week Tracker").Select ' Move to "Next Week Tracker" sheet
    lrtarget = Cells.Find("*", Cells(1, 1), xlFormulas, xlPart, xlByRows, xlPrevious, False).Row ' Find the last row on the sheet
    Cells(lrtarget + 1, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False ' Paste as Values on next row down
    Range("A1").Select ' Reset Cursor
    Sheets("Stats").Select ' Reset Sheet
    
    Range("A1").Select ' Reset Cursor
    Application.ScreenUpdating = True ' Reset Screen Updating

End Sub

Sub JustRefresh()

    ActiveWorkbook.RefreshAll ' Refresh PowerQueries
    Range("A1").Select ' Reset Cursor

End Sub
