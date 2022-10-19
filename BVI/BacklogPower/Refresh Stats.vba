Sub RefreshStats()
'
' RefreshStats Macro
'

'

' Make sure you're on the right sheet
    Sheets("Stats").Select

' Copy yesterdays numbers up to the right field
    Range("Q4").Copy
    Range("Q3").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("R4").Copy
    Range("R3").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
' Copy yesterdays numbers up to the right field
    Range("Q7").Copy
    Range("Q6").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("R7").Copy
    Range("R6").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

' Change to todays date
    Range("P2").FormulaR1C1 = "=TODAY()"
    Range("P2").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
' Refresh PowerQuerys
    ActiveWorkbook.RefreshAll
    
' Copy data to relevant tab (This Weeks Tracker)
    Range("M23:Q23").Copy
    Sheets("This Week Tracker").Select
    lrtarget = Cells.Find("*", Cells(1, 1), xlFormulas, xlPart, xlByRows, xlPrevious, False).Row
    Cells(lrtarget + 1, 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1").Select
    Sheets("Stats").Select

' Copy data to relevant tab (Daily Tracker)
    Range("M26:Q26").Copy
    Sheets("Daily Tracker").Select
    lrtarget = Cells.Find("*", Cells(1, 1), xlFormulas, xlPart, xlByRows, xlPrevious, False).Row
    Cells(lrtarget + 1, 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1").Select
    Sheets("Stats").Select

' Copy data to relevant tab (Next Weeks Tracker)
    Range("M29:Q29").Copy
    Sheets("Next Week Tracker").Select
    lrtarget = Cells.Find("*", Cells(1, 1), xlFormulas, xlPart, xlByRows, xlPrevious, False).Row
    Cells(lrtarget + 1, 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1").Select
    Sheets("Stats").Select
    
    Range("A1").Select

End Sub

