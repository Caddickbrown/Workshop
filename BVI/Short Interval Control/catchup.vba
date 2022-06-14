Sub SIC_Catchup()
    Application.Calculation = xlManual
   Application.ScreenUpdating = False
   Application.DisplayStatusBar = False
   Application.EnableEvents = False
   
    ' before i learnt that you could hve more than 1 variable on a line, defines and sets all the variables needed
    Dim Home As Workbook
    Set Home = ThisWorkbook
    Dim temp As Worksheet
    Set temp = Worksheets("Template")
    Dim SIC As Worksheet
    Dim data As Worksheet
        
    Dim xWorkbook As Workbook
    Dim sht As Worksheet
    Dim Today As Date, start As Date, step As Date
    Today = Date
    Today = Today - 1
    Dim yesterday As Date
    yesterday = Today - 1
    Dim test As String
    Dim begin As Long, final As Long
    Dim SheetName As String
    SheetName = "OverviewInventoryTransactionHis"
    Dim colum_count As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim bay As Integer
    Dim rowcount As Integer
    Dim CreateTime As Integer
    Dim Create As Integer
    Dim clock
    Dim LastComp
    Dim Picks As Integer
    Dim Picker() As String
    Dim pickers As Integer
    Dim outputrow As Integer
    Dim count As Integer
    Dim N_pick As Integer
    Dim N_Pickers As Double
    Dim M_Pick As Integer
    Dim M_Pickers As Double
    Dim A_pick As Integer
    Dim A_Pickers As Double
    Dim previous As String
    Dim Short As Integer
    Dim Date_present As Boolean
    Dim datacount As Long
    
    ' program starts here
    
    
    ' checks if workbook is read only
    If Home.ReadOnly Then
        i = MsgBox("SIC sheet is opened as Read Only. Please reopen the file and ensure that it is not Read Only before trying to run. Do you wish to continue?", vbYesNo)
        If i = 6 Then MsgBox ("Please save file under a different name."): Exit Sub
        If i = 7 Then Home.Close
    End If
    
    ' looks through each workbook that is open
    For Each xWorkbook In Application.Workbooks
        If xWorkbook.Name <> Home.Name Then ' ignores home workbook
            For Each sht In xWorkbook.Worksheets 'looks at each sheet in current workbook
                If sht.Name = SheetName Then
                    column_count = sht.Cells(1, Columns.count).End(xlToLeft).Column
                    For i = 1 To column_count
                        If sht.Cells(1, i) = "Bay" Then bay = i
                        If sht.Cells(1, i) = "Created" Then Create = i
                        If bay > 0 And Create > 0 Then Exit For ' searches row 1 for the column containing the required data if present exits loop having recorded column
                    Next i
                    If sht.Cells(2, bay) = "SOM" Or sht.Cells(2, bay) = "MSOM" Or sht.Cells(2, bay) = "PK" Then located = True: Exit For ' searches the identified row for what needs the required info
                End If
            Next sht
            If located = True Then Exit For
        End If
    Next xWorkbook
    'cancels out of the program if there is no open data file.
        If located = False Then MsgBox ("You must download the data from IFS then rerun the program"): Exit Sub
    'copys in teh data
        xWorkbook.Sheets(SheetName).Copy Before:=Home.Worksheets("Targets")
        Set data = Sheets(SheetName)
    'closes the original IFS download
        xWorkbook.Close SaveChanges:=False
    'counts the rows and columns i the IFS data
    rowcount = data.UsedRange.Rows.count
    column_count = data.Cells(1, Columns.count).End(xlToLeft).Column
    
    'finds time& performed by column
        For i = 1 To column_count
            If data.Cells(1, i) = "Creation Time" Then CreateTime = i
            If data.Cells(1, i) = "Performed By" Then pickers = i
            If CreateTime > 0 And pickers > 0 Then Exit For
        Next i
        
        'sorts by date
        rowcount = data.Cells(Rows.count, Create).End(xlUp).Row
        Range(data.Cells(1, 1), data.Cells(rowcount, column_count)).Sort Key1:=Range(data.Cells(1, Create), data.Cells(rowcount, Create)), Order1:=xlAscending, Header:=xlYes
    
    start = data.Cells(2, Create) ' selects the oldest date
    For step = start To Today 'cycles through fromt eh oldest date until last complete day
        
        test = Format(step, "ddmmmyy") 'puts date in correct format for sheet name
    ' checks if current day has a sheet
        For Each sht In Application.ThisWorkbook.Worksheets
            If sht.Name = test Then Set SIC = sht: GoTo jump 'steps through all workbooks if the required sheet name is present leaves it there, if not creates a new one
        Next sht
        Worksheets("template").Copy After:=Sheets(Sheets.count) 'copies the template
        Set SIC = Sheets(Sheets.count) 'sets teh sheet to use
        SIC.Name = test ' sets the sheet name
        SIC.Cells(1, 13) = step ' copies teh date into the date location
jump:
        begin = final 'starts where the previous loop finished ---- i don't actually think this line does anything.....
        final = 0 ' resets final - don't know why i need to as i reset it agin immediately
        located = False
        For i = 2 To rowcount 'finds the start and end points in teh date set for the current date
            If data.Cells(i, Create) = step And located = False Then begin = i: located = True
            If data.Cells(i, Create) > step Then final = i - 1: Exit For
        Next i
        If final = 0 Then final = i - 1
        
        'sorts by time for selected date
        Range(data.Cells(begin - 1, 1), data.Cells(final, column_count)).Sort Key1:=Range(data.Cells(begin - 1, CreateTime), data.Cells(final, CreateTime)), Order1:=xlAscending, Header:=xlYes
        'resets the picker array
        ReDim Picker(0)
        
        'works through for each hour and records the number of picks and the number of pickers required in each hour.
        For i = 1 To 24
            If j = 0 Then k = begin Else k = j ' sets start location for the loop each hour, if not 0 sets to where teh previous hour ended
            For j = k To final
                'if created time is in teh correct hour
                If Hour(data.Cells(j, CreateTime)) >= i - 1 Then
                If Hour(data.Cells(j, CreateTime)) < i Then
                    pick = pick + 1 'adds 1 to pick
                    If data.Cells(j, bay) = "PK" Then Short = Short + 1 'adds 1 to shortages count
                    For k = 0 To UBound(Picker()) 'checks if the picker has already picked something this hour
                        If data.Cells(j, pickers) = Picker(k) Then GoTo present
                    Next k
                    ReDim Preserve Picker(UBound(Picker()) + 1) 'if not increases teh length of the array
                    Picker(UBound(Picker())) = data.Cells(j, pickers) ' adds teh pickers name to the array
present:
                Else: Exit For
                End If
                End If
            Next j
            'prints out the data
            SIC.Cells(i + 2, 11) = Sheets("Targets").Cells(6, 2) ' person populating
            SIC.Cells(i + 2, 2) = pick 'picks in hour
            SIC.Cells(i + 2, 4) = UBound(Picker()) 'pickers in hour
            If i = 2 Or i = 5 Or i = 10 Or i = 13 Or i = 18 Or i = 21 Then SIC.Cells(i + 2, 5) = Sheets("Targets").Cells(2, 2) * 0.75 Else SIC.Cells(i + 2, 5) = Sheets("Targets").Cells(2, 2) ' target picks per hour
            If SIC.Cells(i + 2, 4) > 0 Then SIC.Cells(i + 2, 6) = Round(pick / SIC.Cells(i + 2, 4), 2) Else SIC.Cells(i + 2, 6) = 0 ' if picks completed calculates teh pick per person for the hour
            If SIC.Cells(i + 2, 6) < SIC.Cells(i + 2, 5) And SIC.Cells(i + 2, 6) > 0 Then SIC.Cells(i + 2, 6).Interior.ColorIndex = 3 Else If SIC.Cells(i + 2, 6) > 0 Then SIC.Cells(i + 2, 6).Interior.ColorIndex = 4 'colours red or greendepending ont eh target achievement
            SIC.Cells(i + 2, 7) = Short ' prints the number of shortages rectified in teh hour
            'resets hourly variables
            pick = 0
            Short = 0
            ReDim Picker(0)
        Next i ' next hour
        
        'prints last date complete
        LastComp = TimeSerial(i - 1, 0, 0)
        SIC.Cells(8, 14) = LastComp
        
        'yesterdays sheet name in required format
        previous = Format(SIC.Cells(1, 13) - 1, "ddmmmyy")
        
        On Error Resume Next
        
        'adds teh picks achieved from 10-11pm & 11pm - midnight to the nightshift total picks & picking hours
        N_pick = Worksheets(previous).Cells(25, 2) + Worksheets(previous).Cells(26, 2)
        N_Pickers = Worksheets(previous).Cells(25, 4) + Worksheets(previous).Cells(26, 4)
        For i = 3 To SIC.Cells(Rows.count, 4).End(xlUp).Row ' cycles through each hour run and adds the picks and picking hours to teh respective shift N (Night) m (morning) or A (afternoon)
            If i <= 8 Then N_pick = N_pick + SIC.Cells(i, 2): If i = 4 Or i = 7 Then N_Pickers = N_Pickers + SIC.Cells(i, 4) * 0.75 Else N_Pickers = N_Pickers + SIC.Cells(i, 4)
            If i <= 16 And i > 8 Then M_Pick = M_Pick + SIC.Cells(i, 2): If i = 12 Or i = 15 Then M_Pickers = M_Pickers + SIC.Cells(i, 4) * 0.75 Else M_Pickers = M_Pickers + SIC.Cells(i, 4)
            If i <= 24 And i > 16 Then A_pick = A_pick + SIC.Cells(i, 2): If i = 20 Or i = 23 Then A_Pickers = A_Pickers + SIC.Cells(i, 4) * 0.75 Else A_Pickers = A_Pickers + SIC.Cells(i, 4)
        Next i
        'prints the picks, pickers and if picks are more than 0 calculates the picks per person per hour for each shift and overall performance colours it red or green based on target
        SIC.Cells(12, 13) = N_pick
        SIC.Cells(12, 14) = N_Pickers
        If SIC.Cells(12, 14) > 0 Then SIC.Cells(12, 15) = Round(SIC.Cells(12, 13) / SIC.Cells(12, 14), 2)
        If SIC.Cells(12, 15) < Worksheets("Targets").Cells(2, 2) And SIC.Cells(12, 15) > 0 Then SIC.Cells(12, 15).Interior.ColorIndex = 3 Else If SIC.Cells(12, 15) > 0 Then SIC.Cells(12, 15).Interior.ColorIndex = 4
        SIC.Cells(13, 13) = M_Pick
        SIC.Cells(13, 14) = M_Pickers
        If SIC.Cells(13, 14) > 0 Then SIC.Cells(13, 15) = Round(SIC.Cells(13, 13) / SIC.Cells(13, 14), 2)
        If SIC.Cells(13, 15) < Sheets("Targets").Cells(2, 2) And SIC.Cells(13, 15) > 0 Then SIC.Cells(13, 15).Interior.ColorIndex = 3 Else If SIC.Cells(13, 15) > 0 Then SIC.Cells(13, 15).Interior.ColorIndex = 4
        SIC.Cells(14, 13) = A_pick
        SIC.Cells(14, 14) = A_Pickers
        If SIC.Cells(14, 14) > 0 Then SIC.Cells(14, 15) = Round(SIC.Cells(14, 13) / SIC.Cells(14, 14), 2)
        If SIC.Cells(14, 15) < Sheets("Targets").Cells(2, 2) And SIC.Cells(14, 15) > 0 Then SIC.Cells(14, 15).Interior.ColorIndex = 3 Else If SIC.Cells(14, 15) > 0 Then SIC.Cells(14, 15).Interior.ColorIndex = 4
        SIC.Cells(15, 13) = SIC.Cells(12, 13) + SIC.Cells(13, 13) + SIC.Cells(14, 13)
        SIC.Cells(15, 14) = SIC.Cells(12, 14) + SIC.Cells(13, 14) + SIC.Cells(14, 14)
        If SIC.Cells(15, 14) > 0 Then SIC.Cells(15, 15) = Round(SIC.Cells(15, 13) / SIC.Cells(15, 14), 2)
        If SIC.Cells(15, 15) < Sheets("Targets").Cells(2, 2) And SIC.Cells(15, 15) > 0 Then SIC.Cells(15, 15).Interior.ColorIndex = 3 Else If SIC.Cells(15, 15) > 0 Then SIC.Cells(15, 15).Interior.ColorIndex = 4
        
        'resets all outputs
        N_pick = 0
        N_Pickers = 0
        M_Pick = 0
        M_Pickers = 0
        A_pick = 0
        A_Pickers = 0
        
    Next step ' next day to catchup on
    
    Application.DisplayAlerts = False 'turns off pop-ups to ask if ok to delete sheet
        Sheets(SheetName).Delete 'Deletes sheet
        Application.DisplayAlerts = True 'turns pop-ups back on to stop me doing something silly
    SIC.Select
    Home.Save
    
    
     Application.EnableEvents = True
   Application.DisplayStatusBar = True
   Application.ScreenUpdating = True
   Application.Calculation = xlAutomatic
End Sub
