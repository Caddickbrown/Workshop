Sub Parts_List()
    
'Dim t As Single
't = Timer
    
    'all relevant sheets needed
    Dim home As Workbook, Need As Worksheet, Kanbans As Worksheet, land As Worksheet, FR As Worksheet
    Set home = ThisWorkbook
    Set Need = Worksheets("Parts Needed")
    Set Kanbans = Worksheets("Kanbans")
    Set land = Worksheets("Coversheet")
    Set FR = Worksheets("FastRunners")
    
    'all variables needed
    Dim PartsWant As Integer, KBact As Long, FRact As Long, kbLoc As Boolean, TotalAct As Long, PartsFound As Long
        
    'Clears the area to put the list
    Range(FR.Cells(2, 1), FR.Cells(FR.UsedRange.Rows.Count, 3)).Clear
        
    'Finds the total number of interactions from the pivot
    TotalAct = Need.Cells(Need.UsedRange.Rows.Count, 2)
    'take from coversheet
    PartsWant = land.Cells(1, 2)
    
    PartsFound = 0
    NeedStep = 3 ' where to start looking down the kits parts needed pivot table
    
    Do While PartsFound < PartsWant And NeedStep < Need.UsedRange.Rows.Count ' less than the parts wanted in FR locations and
        For k = 2 To Kanbans.UsedRange.Rows.Count ' checks if the part number is in a kanban
            If Need.Cells(NeedStep, 1) = Kanbans.Cells(k, 1) Then kbLoc = True: Exit For
        Next k
        If kbLoc = True Then
            KBact = KBact + Need.Cells(NeedStep, 2) ' if in kanban sums total knban interactions
        Else
            PartsFound = PartsFound + 1 ' tracs teh parts found
            FRact = FRact + Need.Cells(NeedStep, 2) ' adds fastrunning interations together
            FR.Cells(PartsFound + 1, 1) = Need.Cells(NeedStep, 1) 'prints the part, interactions, qty
            FR.Cells(PartsFound + 1, 2) = Need.Cells(NeedStep, 2)
            FR.Cells(PartsFound + 1, 3) = Need.Cells(NeedStep, 3)
        End If
        NeedStep = NeedStep + 1 'iterates down the needed sheet
        kbLoc = False
        'Application.StatusBar = "Fast running Parts " & Format(PartsFound / PartsWant, "0.0%")
    Loop
    FR.Cells(1, 11) = KBact / TotalAct 'prints tottal percentage of kanban interactions
    FR.Cells(2, 11) = FRact / TotalAct 'prints percentage of fast runners interactions
    FR.Cells(3, 11) = (KBact + FRact) / TotalAct 'prints total interactions between kanbans and fast runners
    
    FR.Select
    
    Application.StatusBar = False
        
'MsgBox Format((Timer - t) / 86400, "hh:mm:ss")

End Sub
