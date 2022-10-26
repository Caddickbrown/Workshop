Sub Generate Sterilisation List()

Dim sterilistlocation As String, newsheet As String, shipno As String
Dim wbI As Workbook, wbO As Workbook
Dim wsI As Worksheet, wsO As Worksheet


    shipno = ActiveSheet.Name
    'Application.InputBox("What is your Shipment Number?")

    Set wbI = ThisWorkbook
    shipno = wbI.ActiveSheet.Name
    Set wsI = wbI.Sheets(shipno)

    wsI.Range("T21:AC90").Copy

    sterilistlocation = "S:\Public\AA Kit Boxing Data\AA Kit Boxing Data\"
    newsheet = "BVI KITS " & shipno


    Set wbO = Workbooks.Add

    With wbO
        '~~> Set the relevant sheet to where you want to paste
        Set wsO = wbO.Worksheets(1)

        '~~>. Save the file
        .SaveAs Filename:=sterilistlocation & newsheet, FileFormat:=56

        '~~> Copy the range
        wsI.Range("T21:AC90").Copy

        '~~> Paste it in say Cell A1. Change as applicable
        wsO.Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    End With

End Sub




