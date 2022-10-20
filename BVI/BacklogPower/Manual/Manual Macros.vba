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

    Range("A1").Select ' Reset Cursor
    Application.ScreenUpdating = True ' Reset Screen Updating

End Sub

Sub FillTrackers()

   Application.ScreenUpdating = False ' Cleans View up a bit, so it doesn't jump around

' Make sure you're on the right sheet
    Sheets("Stats").Select ' Reset Sheet

' Copy data to relevant tab (This Week)
    Range("M23:Q23").Copy
    Sheets("Archive").Select
    lrtarget = Cells.Find("*", Cells(1, 1), xlFormulas, xlPart, xlByRows, xlPrevious, False).Row ' Find the last row on the sheet
    Cells(lrtarget + 1, 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False ' Paste as Values
    Range("A1").Select ' Reset Cursor
    Sheets("Stats").Select ' Reset Sheet

' Copy data to relevant tab (Daily)
    Range("N26:Q26").Copy
    Sheets("Archive").Select
    lrtarget = Cells.Find("*", Cells(1, 1), xlFormulas, xlPart, xlByRows, xlPrevious, False).Row ' Find the last row on the sheet
    Cells(lrtarget, 6).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False ' Paste as Values
    Range("A1").Select ' Reset Cursor
    Sheets("Stats").Select ' Reset Sheet

' Copy data to relevant tab (Next Week)
    Range("N29:Q29").Copy
    Sheets("Archive").Select
    lrtarget = Cells.Find("*", Cells(1, 1), xlFormulas, xlPart, xlByRows, xlPrevious, False).Row ' Find the last row on the sheet
    Cells(lrtarget, 10).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False ' Paste as Values
    Range("A1").Select ' Reset Cursor
    Sheets("Stats").Select ' Reset Sheet
    
    Range("A1").Select ' Reset Cursor
    Application.ScreenUpdating = True ' Reset Screen Updating

End Sub
