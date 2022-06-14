Sub Today()
    Dim Home As Workbook, SIC As Workbook, xWorkbook As Workbook
    Dim Past As Worksheet, sht As Worksheet
    
    Set Home = ThisWorkbook
    Set Past = Worksheets("Past_Data")
    
    Dim Past_count As Long, outcol As Integer
    Past_count = Past.UsedRange.Rows.Count + 1 ' finds where to start the data from
               
    For Each xWorkbook In Application.Workbooks ' searches all open workbooks to find the short interval control list
        If xWorkbook.Name = "Short_Interval_Control_sheet(SIC).xlsm" Then Set SIC = xWorkbook: Exit For
    Next xWorkbook
    
    SIC.Activate 'activates SIC workbook
    For Each sht In SIC.Worksheets 'steps through each sheet in workbook
    sht.Activate 'activates teh sheet
    
        If sht.Name Like ("##***##") Then 'checks if the sheet name is right in DDMMMYY format
            If sht.Cells(1, 13) > Past.Cells(1, 19) Then 'if the date on teh sheet is greater than teh last completed day
                If sht.Cells(26, 2) Like "#*" Then 'if picks achieved in final hour of day is complete
                    If sht.Cells(15, 13) > 0 Then 'if any picks have been managed in teh day
                        Past.Cells(Past_count, 1) = sht.Cells(1, 13) 'copies the day
                        Past.Cells(Past_count, 2) = Application.WorksheetFunction.IsoWeekNum(Past.Cells(Past_count, 1)) ' finds teh week number
                        outcol = 3
                        For i = 0 To 3 'copies teh 4x3 table of picks, hours& pph into the past data sheet
                            For j = 0 To 2
                                Past.Cells(Past_count, outcol) = sht.Cells(12 + i, 13 + j)
                                outcol = outcol + 1
                            Next j
                        Next i
                        Past.Cells(1, 19) = Past.Cells(Past_count, 1) ' pastes teh last completed date into last comp cell
                        If Past.Cells(Past_count, 2) Mod 2 = 0 And Past.Cells(Past_count, 11) <> "" Then 'colours either blue or red depending on week/shift
                            For i = 6 To 8
                                Past.Cells(Past_count, i).Interior.ColorIndex = 3 ' red
                            Next i
                            For i = 9 To 11
                                Past.Cells(Past_count, i).Interior.ColorIndex = 33 ' blue
                            Next i
                        ElseIf Past.Cells(Past_count, 2) Mod 2 = 1 And Past.Cells(Past_count, 11) <> "" Then
                            For i = 6 To 8
                                Past.Cells(Past_count, i).Interior.ColorIndex = 33
                            Next i
                            For i = 9 To 11
                                Past.Cells(Past_count, i).Interior.ColorIndex = 3
                            Next i
                        End If
                        Past_count = Past_count + 1 'move to next output line
                    End If
                End If
            End If
        End If
                
    Next sht
    Home.Activate 'reactivates the SIC.archive
    
    Call shiftanalysis
    Call WeekNum

End Sub