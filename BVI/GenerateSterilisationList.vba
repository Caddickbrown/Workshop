'ToDo
'- [ ] Will need adding into BID072
'- [ ] Change sterilistlocation Location
'- [ ] Adapt for Malosa
'- [ ] Do we want to automatically save the file?
'- [ ] Extend the ranges?

Sub Generate_Sterilisation_List()

'Setup
    Dim sterilistlocation As String, newsheet As String, shipno As String
    Dim wbI As Workbook, wbO As Workbook
    Dim wsI As Worksheet, wsO As Worksheet
'More Setup
    Set wbI = ThisWorkbook
    shipno = wbI.ActiveSheet.Name
    Set wsI = wbI.Sheets(shipno)

    sterilistlocation = "S:\Public\AA Kit Boxing Data\AA Kit Boxing Data\" 'This needs changing to actual location
    newsheet = "BVI KITS " & shipno 'The name of the new sheet

    Set wbO = Workbooks.Add 'Generate a new workbook and name a variable "wbO"
    Set wsO = wbO.Worksheets(1) 'Set the relevant sheet to where you want to paste

'Save the file as xlsm
    wbO.SaveAs Filename:=sterilistlocation & newsheet, FileFormat:=52

'Copy Data
    wsI.Range("T21:AC90").Copy

'Paste it in say Cell A1. Change as applicable
    wsO.Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

'Header Data
    Range("A1").Value = "BVI " & shipno 'Name of the sheet? Not really sure what this is for
    wsO.Range("B1").Value = wsI.Range("N4").Value 'Copy over P-Number (Purchase Order Number?)
    Range("J3").Formula = "=SUM(J4:J71)" 'Total

'Formatting
    Range("A3:J3").Font.Bold = True 'Bold Row

    With Range("A1:J71").Font 'Font Standardisation
        .Name = "Calibri"
        .Size = 16
    End With

    Range("A1:J71").Borders.LineStyle = xlDouble 'Double Borders

'Colours
'Green
    Range("A1").Interior.Color = 5296274
    Range("C1").Interior.Color = 5296274

'Yellow
    Range("B1").Interior.Color = 65535
    Range("J3").Interior.Color = 65535

'LGrey
    Range("A2:J2").Interior.TintAndShade = -0.149998474074526

'DGrey
    Range("A3:J3").Interior.TintAndShade = -0.249977111117893

'Justify
    With Columns("A:J")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    Range("A1").Select 'Reset Cursor

    Cells.EntireColumn.AutoFit



End Sub




