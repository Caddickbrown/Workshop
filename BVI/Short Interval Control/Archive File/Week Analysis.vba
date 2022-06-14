Sub WeekNum()
    Dim out As Worksheet, Week As Worksheet
    Set out = Worksheets("Past_Data")
    Set Week = Worksheets("WeekNo")
    Dim outcount As Long
    Dim weekcount As Long, StartLoc As Long
    Dim day As String
    Dim daycode As Integer, shift1 As Integer, shift2 As Integer
    
    
    Dim N_pick As Long, N_hour As Double, M_Pick As Long, M_hour As Double, A_pick As Long, A_Hour As Double, WK_Pick As Long, WK_hour As Double
       
    outcount = out.UsedRange.Rows.Count - 1
    weekcount = Week.UsedRange.Rows.Count
    weekstart = out.Cells(Rows.Count, 17).End(xlUp).Row
        
    If Week.Cells(weekcount, 1) = Week.Cells(1, 18) Then weekcount = weekcount + 1
    ' finds which week was last completed
    For i = 3 To outcount
        If out.Cells(i, 16) = Blank Then If Month(out.Cells(i, 1)) = 1 And out.Cells(i, 2) = 52 Then out.Cells(i, 16) = out.Cells(i, 2) + (Year(out.Cells(i, 1)) - 2021 - 1) * 52 Else out.Cells(i, 16) = out.Cells(i, 2) + (Year(out.Cells(i, 1)) - 2021) * 52
        If out.Cells(i, 16) > Week.Cells(1, 18) Then StartLoc = i: Exit For
    Next i
    
    'sums up the total number of picks & hours achieved each week and generates a weekly pph
    For i = StartLoc To outcount
        daycode = Weekday(out.Cells(i, 1), 1)
        If StartLoc = i Then ' for first point of data add
            If 1 = daycode Or daycode = 7 Then 'if sunday or saturday then add to weekend numbers
                    WK_Pick = WK_Pick + out.Cells(i, 12)
                    WK_hour = WK_hour + out.Cells(i, 13)
                Else 'if weekday add to corresponding section night, morning or afternoon
                    N_pick = N_pick + out.Cells(i, 3)
                    N_hour = N_hour + out.Cells(i, 4)
                    M_Pick = M_Pick + out.Cells(i, 6)
                    M_hour = M_hour + out.Cells(i, 7)
                    A_pick = A_pick + out.Cells(i, 9)
                    A_Hour = A_Hour + out.Cells(i, 10)
                End If
        Else
            If out.Cells(i, 2) = out.Cells(i - 1, 2) Then 'if weeknumber is the same as the previous line sum
                If 1 = daycode Or daycode = 7 Then
                    WK_Pick = WK_Pick + out.Cells(i, 12)
                    WK_hour = WK_hour + out.Cells(i, 13)
                Else
                    N_pick = N_pick + out.Cells(i, 3)
                    N_hour = N_hour + out.Cells(i, 4)
                    M_Pick = M_Pick + out.Cells(i, 6)
                    M_hour = M_hour + out.Cells(i, 7)
                    A_pick = A_pick + out.Cells(i, 9)
                    A_Hour = A_Hour + out.Cells(i, 10)
                End If
            Else '
                Week.Cells(weekcount, 1) = out.Cells(i - 1, 2)
                If Month(out.Cells(i - 1, 1)) = 1 And out.Cells(i - 1, 2) = 52 Then Week.Cells(weekcount, 17) = out.Cells(i - 1, 2) + (Year(out.Cells(i - 1, 1)) - 2021 - 1) * 52 Else Week.Cells(weekcount, 17) = out.Cells(i - 1, 2) + (Year(out.Cells(i - 1, 1)) - 2021) * 52
                
                Week.Cells(weekcount, 2) = N_pick
                Week.Cells(weekcount, 3) = N_hour
                If N_hour > 0 Then Week.Cells(weekcount, 4) = Round(N_pick / N_hour, 2) Else Week.Cells(weekcount, 4) = 0
                
                If Week.Cells(weekcount, 1) Mod 2 = 0 Then shift1 = 5: shift2 = 8 Else shift1 = 8: shift2 = 5
                
                Week.Cells(weekcount, shift1) = M_Pick
                Week.Cells(weekcount, shift1 + 1) = M_hour
                If M_hour > 0 Then Week.Cells(weekcount, shift1 + 2) = Round(M_Pick / M_hour, shift1 + 2) Else Week.Cells(weekcount, shift1 + 2) = 0
                
                Week.Cells(weekcount, shift2) = A_pick
                Week.Cells(weekcount, shift2 + 1) = A_Hour
                If A_Hour > 0 Then Week.Cells(weekcount, shift2 + 2) = Round(A_pick / A_Hour, shift2 + 2) Else Week.Cells(weekcount, shift2 + 2) = 0
                Week.Cells(weekcount, 11) = WK_Pick
                Week.Cells(weekcount, 12) = WK_hour
                If WK_hour > 0 Then Week.Cells(weekcount, 13) = Round(WK_Pick / WK_hour, 13) Else Week.Cells(weekcount, 13) = 0
                On Error Resume Next
                Week.Cells(weekcount, 14) = N_pick + M_Pick + A_pick + WK_Pick
                Week.Cells(weekcount, 15) = N_hour + M_hour + A_Hour + WK_hour
                Week.Cells(weekcount, 16) = (N_pick + M_Pick + A_pick + WK_Pick) / (N_hour + M_hour + A_Hour + WK_hour)
                weekcount = weekcount + 1
                On Error GoTo 0
                If 1 = daycode Or daycode = 7 Then
                    N_pick = 0
                    N_hour = 0
                    M_Pick = 0
                    M_hour = 0
                    A_pick = 0
                    A_Hour = 0
                    WK_Pick = out.Cells(i, 12)
                    WK_hour = out.Cells(i, 13)
                Else
                    N_pick = out.Cells(i, 3)
                    N_hour = out.Cells(i, 4)
                    M_Pick = out.Cells(i, 6)
                    M_hour = out.Cells(i, 7)
                    A_pick = out.Cells(i, 9)
                    A_Hour = out.Cells(i, 10)
                    WK_Pick = 0
                    WK_hour = 0
                End If
            End If
        End If
    Next i
    Week.Cells(weekcount, 1) = out.Cells(i - 1, 2)
    Week.Cells(weekcount, 2) = N_pick
    Week.Cells(weekcount, 3) = N_hour
    If N_hour > 0 Then Week.Cells(weekcount, 4) = Round(N_pick / N_hour, 2) Else Week.Cells(weekcount, 4) = 0
    
    If Week.Cells(weekcount, 1) Mod 2 = 0 Then shift1 = 5: shift2 = 8 Else shift1 = 8: shift2 = 5
    If Week.Cells(weekcount, 17) <> "" Then Week.Cells(1, 18) = Week.Cells(weekcount, 17) - 1 Else Week.Cells(1, 18) = Week.Cells(weekcount - 1, 17)
    Week.Cells(weekcount, shift1) = M_Pick
    Week.Cells(weekcount, shift1 + 1) = M_hour
    If M_hour > 0 Then Week.Cells(weekcount, shift1 + 2) = Round(M_Pick / M_hour, shift1 + 2) Else Week.Cells(weekcount, shift1 + 2) = 0
    
    Week.Cells(weekcount, shift2) = A_pick
    Week.Cells(weekcount, shift2 + 1) = A_Hour
    If A_Hour > 0 Then Week.Cells(weekcount, shift2 + 2) = Round(A_pick / A_Hour, shift2 + 2) Else Week.Cells(weekcount, shift2 + 2) = 0
    Week.Cells(weekcount, 11) = WK_Pick
    Week.Cells(weekcount, 12) = WK_hour
    If WK_hour > 0 Then Week.Cells(weekcount, 13) = Round(WK_Pick / WK_hour, 13) Else Week.Cells(weekcount, 13) = 0
    On Error Resume Next
    Week.Cells(weekcount, 14) = N_pick + M_Pick + A_pick + WK_Pick
    Week.Cells(weekcount, 15) = N_hour + M_hour + A_Hour + WK_hour
    Week.Cells(weekcount, 16) = (N_pick + M_Pick + A_pick + WK_Pick) / (N_hour + M_hour + A_Hour + WK_hour)
    
    
    
End Sub