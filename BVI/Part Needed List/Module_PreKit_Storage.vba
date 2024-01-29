Sub PREKIT_Storage()
    'Dim t As Single
    't = Timer

    'all relevant sheets needed
    Dim home As Workbook, Need As Worksheet, kbLoc As Boolean, totKB As Integer, PT As PivotTable, partNo As Integer, dataout As Worksheet, Kanbans As Worksheet, PivOut As Worksheet, land As Worksheet, workdays As Double, required As Double, PK As Worksheet, partStep As Long, TotVolume As Long, kbCount As Integer, BVIweek As Long, MalWeek As Long
    Set home = ThisWorkbook
    Set Need = Worksheets("Parts Needed")
    Set Kanbans = Worksheets("Kanbans")
    Set land = Worksheets("Coversheet")
    Set PK = Worksheets("Pre-Kit Storage")
    Set PivOut = Worksheets("PivotOut")
    
    'finds the sheet of all parts in teh warehouse
    Call LocateSheet("OverviewInventoryPartInStock", "Part No", "Warehouse", dataout, home, "PRE-KIT", "BLANK", "BLANK")
            
    'sorts the needed pivot table into the order by part number for ease of looking at the final sheet.
    Set PT = Need.PivotTables(1) 'selects the first pivot table of the Parts Needed sheet as teh pivot to look at
    PT.PivotFields("Component").AutoSort Order:=xlAscending, Field:="Component" ' sorts by component number
    
    ' finds where the part No and Location no columns are int eh data out sheet.
    For i = 1 To dataout.UsedRange.Columns.Count
        If dataout.Cells(1, i) = "Part No" Then partNo = i: If partNo > 0 And locno > 0 Then Exit For
        If dataout.Cells(1, i) = "Location No" Then locno = i: If partNo > 0 And locno > 0 Then Exit For
    Next i
        
    'clears the prekit storage sheet
    Range(PK.Cells(3, 1), PK.Cells(PK.UsedRange.Rows.Count, 15)).Clear
    
    'sets start point for output locations
    RemRow = 3
    pkrow = 3
    
    'gets teh target volumes from teh coversheet
    BVIweek = land.Cells(3, 2)
    MalWeek = land.Cells(4, 2)
    TotVolume = PivOut.Cells(PivOut.UsedRange.Rows.Count, 3)
    workdays = land.Cells(5, 2)
    
    'calculates the number of days production in teh current plan, this calculates the numbers of times needed for at least 1 use per day.
    required = Application.WorksheetFunction.RoundUp((TotVolume / (BVIweek + MalWeek)) * workdays, 0)
    totKB = Kanbans.UsedRange.Rows.Count - 1
    
    'steps through all the parts in teh parts needed list
    For partStep = 3 To Need.UsedRange.Rows.Count - 1
        If Need.Cells(partStep, 2) > required Then ' if the number of interactions is greater than 1 per day
            If kbCount < totKB Then ' if not all kanban items have been found
                For kbstep = 2 To Kanbans.UsedRange.Rows.Count
                    If Need.Cells(partStep, 1) = Kanbans.Cells(kbstep, 1) Then ' checks if the item is in a kanban
                        kbCount = kbCount + 1
                        kbLoc = True
                        PK.Cells(pkrow, 1) = Need.Cells(partStep, 1) 'prints the part number
                        PK.Cells(pkrow, 2) = "ECA-BULK" ' prints ECA bulk
                        PK.Cells(pkrow, 1).Interior.ColorIndex = 4 'colours green
                        PK.Cells(pkrow, 2).Interior.ColorIndex = 4
                        pkrow = pkrow + 1 ' moves locaiotn down 1
                        Exit For
                    End If
                Next kbstep
            End If
            If kbLoc = False Then ' if it's not in a kanban
                PK.Cells(pkrow, 1) = Need.Cells(partStep, 1) ' prints part number
                For i = 2 To dataout.UsedRange.Rows.Count ' cycles thropugh the parts in storage in prekit
                    If PK.Cells(pkrow, 1) = dataout.Cells(i, partNo) Then ' if the part number needed matches somethin in teh storage location
                        If dataout.Cells(i, locno) Like ("PK-T#*") Then 'if its in a pk-T123 it will pritn the specified numbered tub
                            PK.Cells(pkrow, 2) = dataout.Cells(i, locno)    'tub number
                            PK.Cells(pkrow, 1).Interior.ColorIndex = 3      'colours red
                            PK.Cells(pkrow, 2).Interior.ColorIndex = 3
                        ElseIf dataout.Cells(i, locno) Like ("ECA*") Then
                            PK.Cells(pkrow, 2) = dataout.Cells(i, locno)    'prints eca bulk
                            PK.Cells(pkrow, 1).Interior.ColorIndex = 4      ' colours green
                            PK.Cells(pkrow, 2).Interior.ColorIndex = 4
                        ElseIf dataout.Cells(i, locno) Like ("PKMAL*") Then 'if pk malosa tub
                            PK.Cells(pkrow, 2) = dataout.Cells(i, locno)    'prints tub reference
                            PK.Cells(pkrow, 1).Interior.ColorIndex = 26     'colours purple
                            PK.Cells(pkrow, 2).Interior.ColorIndex = 26
                        ElseIf dataout.Cells(i, locno) Like ("PK-S#*") Then  ' pk-shelf generally used for drapes
                            PK.Cells(pkrow, 2) = dataout.Cells(i, locno)     ' shelf reference
                            PK.Cells(pkrow, 1).Interior.ColorIndex = 6      ' colours yellow
                            PK.Cells(pkrow, 2).Interior.ColorIndex = 6
                        ElseIf dataout.Cells(i, locno) Like ("*SHORT*") Or dataout.Cells(i, locno) Like ("*KB*") Or dataout.Cells(i, locno) Like ("*TEMP*") Then 'other none defined locatons
                            PK.Cells(pkrow, 2) = "PK-TEMP"                  'prints PK-temp
                            PK.Cells(pkrow, 1).Interior.ColorIndex = 33     'colours blue
                            PK.Cells(pkrow, 2).Interior.ColorIndex = 33
                        Else    'if not on the list
                            PK.Cells(pkrow, 2) = "PK-TEMP"                  'prints Pk-temp
                            PK.Cells(pkrow, 1).Interior.ColorIndex = 33     'colours blue
                            PK.Cells(pkrow, 2).Interior.ColorIndex = 33
                        End If
                        Exit For
                    End If
                    
                Next i
                If PK.Cells(pkrow, 2) = "" Then PK.Cells(pkrow, 2) = "PK-TEMP": PK.Cells(pkrow, 1).Interior.ColorIndex = 33: PK.Cells(pkrow, 2).Interior.ColorIndex = 33 ' if blank pk-tem p& blue
                pkrow = pkrow + 1 ' adds 1 to the pkro printed
            End If
        ElseIf Need.Cells(partStep, 2) >= required / 2 Then 'If the number of interactions is greater than or equal to 1/2 th erequired rate
            For i = 2 To dataout.UsedRange.Rows.Count 'cycles through parts used
                If Need.Cells(partStep, 1) = dataout.Cells(i, partNo) Then ' finds part if on current list
                    If dataout.Cells(i, locno) Like ("PK-T#*") Then 'sets & colours each locaiton either PK-T123 PKMAL, PK - shelf or ECA. If will not be stored if it doesnot already have a locaiton.
                        PK.Cells(pkrow, 1) = dataout.Cells(i, partNo)
                        PK.Cells(pkrow, 2) = dataout.Cells(i, locno)
                        PK.Cells(pkrow, 1).Interior.ColorIndex = 3
                        PK.Cells(pkrow, 2).Interior.ColorIndex = 3
                        pkrow = pkrow + 1
                    ElseIf dataout.Cells(i, locno) Like ("ECA*") Then
                        PK.Cells(pkrow, 1) = dataout.Cells(i, partNo)
                        PK.Cells(pkrow, 2) = dataout.Cells(i, locno)
                        PK.Cells(pkrow, 1).Interior.ColorIndex = 4
                        PK.Cells(pkrow, 2).Interior.ColorIndex = 4
                        pkrow = pkrow + 1
                    ElseIf dataout.Cells(i, locno) Like ("PKMAL*") Then
                        PK.Cells(pkrow, 1) = dataout.Cells(i, partNo)
                        PK.Cells(pkrow, 2) = dataout.Cells(i, locno)
                        PK.Cells(pkrow, 1).Interior.ColorIndex = 26
                        PK.Cells(pkrow, 2).Interior.ColorIndex = 26
                        pkrow = pkrow + 1
                    ElseIf dataout.Cells(i, locno) Like ("PK-S#*") Then
                        PK.Cells(pkrow, 1) = dataout.Cells(i, partNo)
                        PK.Cells(pkrow, 2) = dataout.Cells(i, locno)
                        PK.Cells(pkrow, 1).Interior.ColorIndex = 6
                        PK.Cells(pkrow, 2).Interior.ColorIndex = 6
                        pkrow = pkrow + 1
                    End If
                    Exit For
                End If
            Next i
        ElseIf Need.Cells(partStep, 2) < required / 2 Then 'if it needed less than every other day
            For i = 2 To dataout.UsedRange.Rows.Count
                If Need.Cells(partStep, 1) = dataout.Cells(i, partNo) Then ' copies the part number into a list to be removed
                    If dataout.Cells(i, locno) Like ("PK-T#*") Or dataout.Cells(i, locno) Like ("PK-S#*") Or dataout.Cells(i, locno) Like ("PKMAL*") Or dataout.Cells(i, locno) Like ("ECA*") Then
                        PK.Cells(RemRow, 11) = dataout.Cells(i, partNo)
                        PK.Cells(RemRow, 12) = dataout.Cells(i, locno)
                        RemRow = RemRow + 1
                    End If
                    Exit For
                End If
            Next i
        End If
        kbLoc = False
        'adds part number and locaiton headings at points through the doocument to put them in the right place in word.
        If (pkrow - 2) Mod 24 = 0 Then
            PK.Cells(pkrow, 1) = "Part Number"
            PK.Cells(pkrow, 2) = "Location"
            pkrow = pkrow + 1
        End If
    Next partStep
    
    ' resorts neeeded parts by order of interacitons all other parts need it that way.
    
    PT.PivotFields("Component").AutoSort Order:=xlDescending, Field:="Sum of Interactions"
    
    'Deletes IFS Manufacturing structures
    Application.DisplayAlerts = False 'turns off pop-ups to ask if ok to delete sheet
        dataout.Delete 'Deletes sheet
    Application.DisplayAlerts = True 'turns pop-ups back on to stop me doing something silly
    
    'copies the data
    Range(PK.Cells(2, 1), PK.Cells(PK.Cells(PK.Rows.Count, 1).End(xlUp).Row, 2)).Copy

    Dim wordApp As Word.Application
    Dim doc As Word.Document, folder As String, file As String, list As Table, UpdateDate As String, today As Date
    
    'sets folder path and file name. ENSURE THEY ARE STORED IN THE SAME FOLDER
    folder = home.Path
    file = "\BID250.dotx" 'D250 template opens word and the template
    Set wordApp = CreateObject("word.application")
    Set D250 = wordApp.Documents.Open(folder & file)
    wordApp.Visible = True
    
    'Set D250 = wordApp.Documents.Open(folder & "\D250-07-03-2022.docx")
    'copies teh list into teh template & sets teh font size.
    With D250.Paragraphs(D250.Paragraphs.Count).Range
        .PasteExcelTable _
        LinkedToExcel:=False, _
        WordFormatting:=False, _
        RTF:=True
        .Font.Size = 18
    End With
    
    'adjusts column widths of table
    D250.Tables(1).Columns(1).SetWidth ColumnWidth:=CentimetersToPoints(4.5), RulerStyle:=wdAdjustNone
    D250.Tables(1).Columns(2).SetWidth ColumnWidth:=CentimetersToPoints(4#), RulerStyle:=wdAdjustNone
    'aligns to teh center of the text
    D250.Tables(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    
    'Sets todays date and puts into format desired
    today = Date
    UpdateDate = Format(today, "dd-mm-yyyy")
    
    D250.Select
    
    'finds each bookmark to fill in the date
    For Each bmk In D250.Bookmarks
        bmk.Range.Text = UpdateDate
    Next bmk
    
    'replaces the 2 part numbers which loose 000 at teh start with the correct part numbers.
    With D250.Range.Find
        .Text = "8681"
        .Replacement.Text = "0008681"
        .Execute Replace:=wdReplaceAll
        .Text = "8685"
        .Replacement.Text = "0008685"
        .Execute Replace:=wdReplaceAll
    End With
    
    'Saves the exported sheet
    D250.SaveAs Filename:=folder & "\BID250-" & UpdateDate, FileFormat:=wdFormatDocumentDefault
       
End Sub
