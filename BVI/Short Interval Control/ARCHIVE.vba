Sub Archive()
    'bits to speed up macro
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
   
    Dim Home As Workbook
    Set Home = ThisWorkbook
    
    ' counts the number of open sheets in the SIC workbook
    Dim count As Integer
    count = Home.Sheets.count ' count = number of sheets in workbook
    
    'Variables for use
    Dim sht As Worksheet
    Dim oldest As Date
    Dim newest As Date
    Dim days As Integer
    Dim present As Boolean
    
    'defines the location and name of the file to store in
    Dim folder As String
    folder = ThisWorkbook.Path ' as long as archive and final SIC sheet are stored in the same folder this will work.
    Dim file As String
    file = "SIC_ARCHIVE.xlsm" ' file to store data in
    Dim Archive As Workbook
    Dim filepath As String
    filepath = folder & "\" & file
    
    'If more sheets than the 3 required sheets (Targets, Instructions, Template) & 5 more sheets investigate and move sheets, otherwise skip the function
    If count > 8 Then
    
        'deletes any blank unnamed sheets and records the number of production days data and start and end dates
        For Each sht In Home.Worksheets
            If sht.Name Like ("Sheet*") Then ' searches and deletes anything that is titled SheetX
                If Application.WorksheetFunction.CountA(sht.Cells) = 0 Then Application.DisplayAlerts = False: sht.Delete: Application.DisplayAlerts = True ' checks if there is a blank sheet that has been created but has no data in it if so delete
                If Sheets.count <= 8 Then Exit Sub ' if still more than 8 sheets continues
            ElseIf sht.Name Like ("##***##") Then ' records number of days worth of production and start and finish date
                days = days + 1
                If oldest = #12:00:00 AM# Then oldest = sht.Cells(1, 13) Else If oldest > sht.Cells(1, 13) Then oldest = sht.Cells(1, 13)
                If newest = #12:00:00 AM# Then newest = sht.Cells(1, 13) Else If newest < sht.Cells(1, 13) Then newest = sht.Cells(1, 13)
            End If
        Next sht
        
        
        If days > 5 Then ' if more than 5 days or production move oldest SIC sheets to archive file
            Workbooks.Open (filepath) ' opens the workbook as defined above
            Set Archive = Workbooks(file) ' stores it as archive workboook
            If Archive.ReadOnly = True Then Exit Sub ' Cancels if the archive has been opened as read only this will result in the program automatically saving a copy to the PC rather than keeping all the data in the archive.
move:
            Home.Worksheets(Format(oldest, "ddmmmyy")).move After:=Archive.Worksheets(Worksheets.count) 'moves the oldest sheet to archive
            days = days - 1 ' subtracts 1 from day count
            If days > 5 Then oldest = oldest + 1: Call exists(oldest, newest, Home, present): If present = True Then GoTo move 'if still more than 5 days worth of production it will find the next oldest date (exists function) and move the sheet to archive
        End If
        Archive.Worksheets("Past_Data").Select
        Archive.Save
        Archive.Close
    End If
    'Archive.Save
    'Archive.Close
    
    Application.EnableEvents = True
    Application.DisplayStatusBar = True
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    
End Sub

Function exists(old As Date, newest As Date, Home As Workbook, present As Boolean)
    
    Dim sht As Worksheet
    present = False
redo:
    For Each sht In Home.Worksheets ' checks the new oldest date to confirm if there is a sheet called it, it steps through all sheets in the workbook and checks the sheet name against the desired date. if present exits the funtion
        If sht.Name = Format(old, "ddmmmyy") Then present = True: Exit Function
    Next sht
    
    If old < (newest - 5) Then old = old + 1: GoTo redo 'if not present adds 1 to the date and goes again

End Function