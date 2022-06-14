Sub SIC()
GoTo skip
     'bits to speed up macro
   Application.Calculation = xlManual
   Application.ScreenUpdating = False
   Application.DisplayStatusBar = False
   Application.EnableEvents = False
skip:

    ' Defines all required workbooks & worksheet variables and assigns the set ones.
    Dim Home As Workbook, temp As Worksheet, SIC As Worksheet, data As Worksheet, xWorkbook As Workbook, sht As Worksheet
    Set Home = ThisWorkbook
    Set temp = Worksheets("Template")
        
    'Defines the days to look at
    Dim Today As Date, yesterday As Date
    Today = Date
    yesterday = Today - 1
    
    ' Defines required variables
    Dim test As String, SheetName As String
    test = Format(Today, "ddmmmyy")
    SheetName = "OverviewInventoryTransactionHis"
    Dim colum_count As Integer, i As Integer, j As Integer, k As Integer, bay As Integer, rowcount As Integer, CreateTime As Integer, Create As Integer
    Dim clock ' no set stype because this needs to be a time which is Date & Time
    Dim LastComp ' no set stype because this needs to be a time which is Date & Time
    Dim Picks As Integer, Picker() As String, pickers As Integer, outputrow As Integer, count As Integer, N_pick As Integer, N_Pickers As Double, M_Pick As Integer, M_Pickers As Double, A_pick As Integer, A_Pickers As Double, previous As String, Short As Integer
    Dim Date_present As Boolean
    
    ' checks if workbook is read only
    If Home.ReadOnly Then
        i = MsgBox("SIC sheet is opened as Read Only. Please reopen the file and ensure that it is not Read Only before trying to run. Do you wisjh to continue?", vbYesNo)
        If i = 6 Then MsgBox ("Please save file under a different name."): Exit Sub
        If i = 7 Then Home.Close
    End If
    ' checks if both today & yesterday have a sheet
    For i = -1 To 0
        test = Format(Today + i, "ddmmmyy") ' formats the date into the required format for the sheet name
        For Each sht In Application.ThisWorkbook.Worksheets 'steps through all sheets i the home workbook to identify if any match the date required
            If sht.Name = test Then Set SIC = sht: GoTo jump ' if the sheet is present set it to be the worksheet required and dont create a new sheet
        Next sht
        'if the sheet is not found, copy the template sheet to the end of the book
        temp.Copy After:=Sheets(Sheets.count)
        'set the SIC worksheet as the last sheet in teh workbook
        Set SIC = Sheets(Sheets.count)
        'rename the SIC sheet with teh required date in correct format
        SIC.Name = test
        ' Put the date into the required cell on the sheet
        SIC.Cells(1, 13) = Today + i
jump:
    Next i
   
   
   
   Call Archive ' calls the archive to move out old sheets if required
     
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
                    If sht.Cells(2, Create) = Today Or sht.Cells(2, Create) = yesterday Then
                        Date_present = True
                        If sht.Cells(2, bay) = "SOM" Or sht.Cells(2, bay) = "MSOM" Or sht.Cells(2, bay) = "PK" Then located = True: Exit For ' searches the identified row for what needs the required info
                    Else: bay = 0: Create = 0
                    End If
                End If
            Next sht
            If located = True Then Exit For
        End If
    Next xWorkbook
    
    'cancels out of the program if there is no open data file.
    If Date_present = False Then MsgBox ("There is no data open from the correct dates, the data should be from " & yesterday & " or " & Today): Exit Sub
    If located = False Then MsgBox ("You must download the data from IFS then rerun the program"): Exit Sub
    
    'copies the data into the home workbook
    xWorkbook.Sheets(SheetName).Copy Before:=Home.Worksheets("Targets")
    Set data = Sheets(SheetName)
    'closes the original data sheet
    xWorkbook.Close SaveChanges:=False
    
    'creates a date in teh required format
    test = Format(data.Cells(2, Create), "ddmmmyy")
    Set SIC = Sheets(test)
    SIC.Name = test
        
    'Finds the length of the data
    column_count = data.Cells(1, Columns.count).End(xlToLeft).Column
    'finds time& performed by columns as everyones IFS put these in slightly different places (it was very annoying)
    For i = 1 To column_count
        If data.Cells(1, i) = "Creation Time" Then CreateTime = i
        If data.Cells(1, i) = "Performed By" Then pickers = i
        If CreateTime > 0 And pickers > 0 Then Exit For
    Next i
    
    'sorts by time
    rowcount = data.Cells(Rows.count, CreateTime).End(xlUp).Row
    'sorts by time
    Range(data.Cells(1, 1), data.Cells(rowcount, column_count)).Sort Key1:=Range(data.Cells(1, CreateTime), data.Cells(rowcount, CreateTime)), Order1:=xlAscending, Header:=xlYes
    
    
    j = 0
    located = False
    'Last complete hour
    clock = Time
    ReDim Picker(0)
    'specifies how far through the day to run the program if today up until the last full hour, if yesterday all day.
    If data.Cells(2, Create) = Today Then count = CInt(Hour(clock)) Else If data.Cells(2, Create) = yesterday Then count = 24
    
    'states where to start running the program to the end
    For i = Hour(SIC.Cells(8, 14)) + 1 To count
        If j = 0 Then k = 2 Else k = j ' sets the start point for data reading if j = 0 then start on line 2 first line of data otherwise set k to equal the final j that was reached last cycle
        For j = k To rowcount
            If Hour(data.Cells(j, CreateTime)) > i - 2 Then 'if the time is between the the previous hour and the current then record
            If Hour(data.Cells(j, CreateTime)) < i Then
                pick = pick + 1 'increase number of picks in teh hour
                If data.Cells(j, bay) = "PK" Then Short = Short + 1
                For k = 0 To UBound(Picker()) ' cycles through teh number of picker who have picked in teh hour if present jumps to end
                    If data.Cells(j, pickers) = Picker(k) Then GoTo present
                Next k
                ReDim Preserve Picker(UBound(Picker()) + 1) 'if not increases the size of array and stores the pickers name
                Picker(UBound(Picker())) = data.Cells(j, pickers)
present:
            Else: Exit For 'if the data is not in the current hour exits th eloop
            End If
            End If
        Next j
        SIC.Cells(i + 2, 11) = Sheets("Targets").Cells(6, 2) ' ignor this this was from when i was having the person who ran it recorded
        SIC.Cells(i + 2, 2) = pick ' prints the number of picks achieved
        SIC.Cells(i + 2, 4) = UBound(Picker()) ' prints the number of pickers who have picked hour
        If i = 2 Or i = 5 Or i = 10 Or i = 13 Or i = 18 Or i = 21 Then SIC.Cells(i + 2, 5) = Sheets("Targets").Cells(2, 2) * 0.75 Else SIC.Cells(i + 2, 5) = Sheets("Targets").Cells(2, 2) ' prints the pick rate required per hour based on when break times are
        If SIC.Cells(i + 2, 4) > 0 Then SIC.Cells(i + 2, 6) = Round(pick / SIC.Cells(i + 2, 4), 2) Else SIC.Cells(i + 2, 6) = 0 ' if picks greater than 0 divides picks by people to get picks per hour else 0
        If SIC.Cells(i + 2, 6) < SIC.Cells(i + 2, 5) And SIC.Cells(i + 2, 6) > 0 Then SIC.Cells(i + 2, 6).Interior.ColorIndex = 3 Else If SIC.Cells(i + 2, 6) > 0 Then SIC.Cells(i + 2, 6).Interior.ColorIndex = 4 ' colors the picks red or green depnding on if its larger or smaller than teh required rate.
        SIC.Cells(i + 2, 7) = Short 'records number of shortages
        'resets all to 0
        pick = 0
        Short = 0
        ReDim Picker(0)
    Next i ' loops round until all hours are run
    
    ' prints last completed serial number
    LastComp = TimeSerial(i - 1, 0, 0)
    SIC.Cells(8, 14) = LastComp
    
    Application.DisplayAlerts = False 'turns off pop-ups to ask if ok to delete sheet
    Sheets(SheetName).Delete 'Deletes sheet
    Application.DisplayAlerts = True 'turns pop-ups back on to stop me doing something silly
    
    previous = Format(SIC.Cells(1, 13) - 1, "ddmmmyy") ' formats the previous days date in teh required format
    
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
    
    'selects the sheet
    SIC.Select
    Home.Save 'saves teh workbook
    
    Application.EnableEvents = True
    Application.DisplayStatusBar = True
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    
End Sub