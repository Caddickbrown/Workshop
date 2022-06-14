Sub History()
    
    'sets & defines varibles and start position
    Dim out As Worksheet, sht As Worksheet
    Set out = Worksheets("Past_Data")
    Dim outrow As Integer, outcol As Integer
    outrow = out.UsedRange.Rows.Count + 1
    Dim i As Integer, j As Integer
    
    
    'steps through each shift and copies the summarised data
    
    For Each sht In ThisWorkbook.Worksheets
        If sht.Name <> "Past_Data" And sht.Name <> "WeekNo" Then
            'checks if the date on the individual sheet is greater than teh last completed, if not moves to next sheet
            If sht.Cells(1, 13) > out.Cells(1, 19) Then
                out.Cells(outrow, 1) = sht.Cells(1, 13) ' pastes teh date
                out.Cells(outrow, 2) = Application.WorksheetFunction.IsoWeekNum(out.Cells(outrow, 1)) ' makes use of isoweek function in excel
                outcol = 3
                'copies the 3 by 4 table with teh daily summary into the sheet offsetting the location point by 1 cell each time by outcol
                For i = 0 To 3
                    For j = 0 To 2
                        out.Cells(outrow, outcol) = sht.Cells(12 + i, 13 + j)
                        outcol = outcol + 1
                    Next j
                Next i
                'pastes teh completed date into teh last completed
                out.Cells(1, 19) = out.Cells(outrow, 1)
                'alternates blue and red shift based on week to simulate the swapped shifts
                If out.Cells(outrow, 2) Mod 2 = 0 And out.Cells(outrow, 11) <> "" Then
                    For i = 6 To 8
                        out.Cells(outrow, i).Interior.ColorIndex = 3 ' RED
                    Next i
                    For i = 9 To 11
                        out.Cells(outrow, i).Interior.ColorIndex = 33 'BLUE
                    Next i
                ElseIf out.Cells(outrow, 2) Mod 2 = 1 And out.Cells(outrow, 11) <> "" Then
                    For i = 6 To 8
                        out.Cells(outrow, i).Interior.ColorIndex = 33
                    Next i
                    For i = 9 To 11
                        out.Cells(outrow, i).Interior.ColorIndex = 3
                    Next i
                End If
                outrow = outrow + 1
            End If
        End If
    Next sht
    
    Call shiftanalysis
    Call WeekNum



End Sub
