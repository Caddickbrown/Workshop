'ToDo
'- [ ] Will need adding into BID072
'- [ ] Change sterilistlocation Location
'- [ ] Adapt for Malosa
'- [ ] Sort Formatting

Sub Generate_Sterilisation_List()

    Dim sterilistlocation As String, newsheet As String, shipno As String
    Dim wbI As Workbook, wbO As Workbook
    Dim wsI As Worksheet, wsO As Worksheet

    Set wbI = ThisWorkbook
    shipno = wbI.ActiveSheet.Name
    Set wsI = wbI.Sheets(shipno)

    wsI.Range("T21:AC90").Copy

    sterilistlocation = "S:\Public\AA Kit Boxing Data\AA Kit Boxing Data\" 'This needs changing to actual location
    newsheet = "BVI KITS " & shipno

    Set wbO = Workbooks.Add

    With wbO
        'Set the relevant sheet to where you want to paste
        Set wsO = wbO.Worksheets(1)

        'Save the file
        .SaveAs Filename:=sterilistlocation & newsheet, FileFormat:=56

        'Copy the range
        wsI.Range("T21:AC90").Copy

        'Paste it in say Cell A1. Change as applicable
        wsO.Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    End With

'Header Data
    Range("A1").Value = "BVI " & shipno
    wsO.Range("B1").Value = wbI.wsI.Range("N4").Value
    Range("J3").FormulaR1C1 = "=SUM(J4:J71)"



'Formatting

    Range("A1:J71").Select

    With Selection.Font
        .Name = "Calibri"
        .Size = 16
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With

    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With

    Range("A3:J3").Font.Bold = True

'Green
    With Range("A1").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

'Yellow
    With Range("B1").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

'LGrey
    With Range("A2:J2").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With

'DGrey
    With Range("A3:J3").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With

End Sub




