Function LocateSheet(SheetName As String, Check1 As String, Check2 As String, dataout As Worksheet, home As Workbook, C2vs1 As String, C2vs2 As String, C2vs3 As String)

'steps through all open workbooks until it finds one with the correct sheet name based on IFS download
Dim Xworkbook As Workbook, sht As Worksheet, C1Loc As Integer, C2Loc As Integer

    For Each Xworkbook In Application.Workbooks ' steps through all open workbooks in excel
        If Xworkbook.Name <> home.Name Then ' ignores home workbook
            For Each sht In Xworkbook.Worksheets 'looks at each sheet in current workbook
                If Trim(Left(sht.Name, InStr(sht.Name, " "))) = SheetName Then
                    column_count = sht.Cells(1, Columns.Count).End(xlToLeft).Column
                    For i = 1 To column_count
                        If sht.Cells(1, i) = Check1 Then C1Loc = i
                        If sht.Cells(1, i) = Check2 Then C2Loc = i
                        If C1Loc > 0 And C2Loc > 0 Then Exit For ' searches row 1 for the column containing the required data if present exits loop having recorded column
                    Next i
                    If sht.Cells(2, C2Loc) = C2vs1 Or sht.Cells(2, C2Loc) = C2vs2 Or sht.Cells(2, C2Loc) = C2vs3 Then
                        located = True
                        Exit For ' searches the identified row for what needs the required info
                    Else: C1Loc = 0: C2Loc = 0 ' resets locations if incorrect
                    End If
                End If
            Next sht
            If located = True Then Exit For
        End If
    Next Xworkbook
    
    'If it hasn''t found the data tells you to download then exits
    If located = False Then MsgBox ("You must download the data from IFS then rerun the program"): Exit Function
    
    'copies teh identified sheet to teh start of the required workbook
    sht.Copy before:=home.Worksheets(1)
    Set dataout = home.Worksheets(1)
    
    'closes the downloaded workbook
    Xworkbook.Close SaveChanges:=False


End Function
