Sub WeekNum2()
    Dim out As Worksheet, Week As Worksheet
    Set out = Worksheets("Past_Data")
    Set Week = Worksheets("WeekNo2")
    Dim outcount As Long
    Dim weekcount As Long, StartLoc As Long
    Dim day As String
    Dim daycode As Integer, shift1 As Integer, shift2 As Integer
    
    
    Dim N_pick As Long, N_hour As Double, M_Pick As Long, M_hour As Double, A_pick As Long, A_Hour As Double, WK_Pick As Long, WK_hour As Double
       
    outcount = out.UsedRange.Rows.Count
    weekcount = Week.UsedRange.Rows.Count
        
    If Week.Cells(weekcount, 1) = Week.Cells(1, 11) Then weekcount = weekcount + 1
    For i = 3 To outcount
        If out.Cells(i, 2) > Week.Cells(1, 11) Then StartLoc = i: Exit For
    Next i
    
    
    
    For i = StartLoc To outcount
        daycode = Weekday(out.Cells(i, 1), 1)
        If StartLoc = i Then
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
        Else
            If out.Cells(i, 2) = out.Cells(i - 1, 2) Then
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
            Else
                Week.Cells(weekcount, 1) = out.Cells(i - 1, 2)
                Week.Cells(1, 11) = out.Cells(i - 1, 2)
                If N_hour > 0 Then Week.Cells(weekcount, 2) = Round(N_pick / N_hour, 2) Else Week.Cells(weekcount, 2) = 0
                If Week.Cells(weekcount, 1) Mod 2 = 0 Then shift1 = 3: shift2 = 4 Else shift1 = 4: shift2 = 3
                If M_hour > 0 Then Week.Cells(weekcount, shift1) = Round(M_Pick / M_hour, shift1) Else Week.Cells(weekcount, shift1) = 0
                If A_Hour > 0 Then Week.Cells(weekcount, shift2) = Round(A_pick / A_Hour, shift2) Else Week.Cells(weekcount, shift2) = 0
                If WK_hour > 0 Then Week.Cells(weekcount, 5) = Round(WK_Pick / WK_hour, 2) Else Week.Cells(weekcount, 5) = 0
                On Error Resume Next
                Week.Cells(weekcount, 6) = (N_pick + M_Pick + A_pick + WK_Pick) / (N_hour + M_hour + A_Hour + WK_hour)
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
    If Week.Cells(weekcount, 1) Mod 2 = 0 Then shift1 = 3: shift2 = 4 Else shift1 = 4: shift2 = 3
    If N_hour > 0 Then Week.Cells(weekcount, 2) = Round(N_pick / N_hour, 2) Else Week.Cells(weekcount, 2) = 0
    If M_hour > 0 Then Week.Cells(weekcount, shift1) = Round(M_Pick / M_hour, shift1) Else Week.Cells(weekcount, shift1) = 0
    If A_Hour > 0 Then Week.Cells(weekcount, shift2) = Round(A_pick / A_Hour, shift2) Else Week.Cells(weekcount, shift2) = 0
    If WK_hour > 0 Then Week.Cells(weekcount, 5) = Round(WK_Pick / WK_hour, 2) Else Week.Cells(weekcount, 5) = 0
    On Error Resume Next
    Week.Cells(weekcount, 6) = (N_pick + M_Pick + A_pick + WK_Pick) / (N_hour + M_hour + A_Hour + WK_hour)
        
End Sub