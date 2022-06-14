Sub shiftanalysis()
'sets & defines varibles and start position
Dim out As Worksheet, Week As Worksheet
    Set out = Worksheets("Past_Data")
    Set shift = Worksheets("by shift")
    Dim outcount As Long
    Dim shiftcount As Long, StartLoc As Long
    Dim day As String
    Dim daycode As Integer, shift1 As Integer, shift2 As Integer
    
    
    Dim N_pick As Long, N_hour As Double, M_Pick As Long, M_hour As Double, A_pick As Long, A_Hour As Double, WK_Pick As Long, WK_hour As Double
       
    'finds the length of the current data sheet and shift sheet
    outcount = out.UsedRange.Rows.Count
    shiftcount = shift.UsedRange.Rows.Count
        
    'places data under the current row not ontop of
    If shift.Cells(shiftcount, 1) = shift.Cells(1, 20) Then shiftcount = shiftcount + 1
    
    ' finds which week was last completed
    For i = 3 To outcount
        test = shift.Cells(1, 20)
        If out.Cells(i, 1) > shift.Cells(1, 20) Then StartLoc = i: Exit For Else StartLoc = i
    Next i
    
    'FOr start point in data list to end copies across relevant data
    For i = StartLoc To outcount
        'finds the day and copies across teh date and week number
        daycode = Weekday(out.Cells(i, 1), 1)
        shift.Cells(i, 1) = out.Cells(i, 1)
        shift.Cells(i, 2) = out.Cells(i, 2)
        'if its a weekened copies to weekend column
        If 1 = daycode Or daycode = 7 Then
            For j = 0 To 2
                shift.Cells(i, 3 + j) = out.Cells(i, 3 + j)
                If j <> 2 Then shift.Cells(i, 12 + j) = out.Cells(i, 6 + j) + out.Cells(i, 9 + j) Else If shift.Cells(i, 13) > 0 Then shift.Cells(i, 14) = Round(shift.Cells(i, 12) / shift.Cells(i, 13), 2)
            Next j
        Else
        'sets location based onshift and copies across
            If shift.Cells(i, 2) Mod 2 = 0 Then red = 6: blue = 9 Else red = 9: blue = 6
            For j = 0 To 2
                shift.Cells(i, 3 + j) = out.Cells(i, 3 + j)
                If j <> 2 Then
                    shift.Cells(i, 6 + j) = out.Cells(i, red + j)
                    shift.Cells(i, 9 + j) = out.Cells(i, blue + j)
                Else
                On Error Resume Next ' DIVIDES the picks/pick hours to get pph
                    shift.Cells(i, 8) = Round(shift.Cells(i, 6) / shift.Cells(i, 7), 2)
                    shift.Cells(i, 11) = Round(shift.Cells(i, 9) / shift.Cells(i, 10), 2)
                On Error GoTo 0
                End If
            Next j
        End If
        'copies total sum
        For j = 0 To 2
            shift.Cells(i, 15 + j) = out.Cells(i, 12 + j)
        Next j
        'pastes last completed date
        shift.Cells(1, 20) = shift.Cells(i, 1)
    Next i
        
End Sub
