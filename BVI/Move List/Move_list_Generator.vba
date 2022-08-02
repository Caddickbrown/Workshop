Sub Move_list_generator()
    ' defines variables
    Dim xworkbook As Workbook, home As Workbook, BVI As Workbook, Malosa As Workbook
    Set home = ThisWorkbook
    Dim sht As Worksheet, list As Worksheet, box_qty As Worksheet, required As Worksheet, M_Required As Worksheet, Cons As Worksheet
    Set list = Worksheets("Amco Pick list")
    Set box_qty = Worksheets("Box Qty")
    Set Cons = Worksheets("Consumables")
    Dim column_count As Integer, boxes As Long, Need As Long, delete As Long, yesno As Integer, outrow As Long
    Dim SheetName As String
    SheetName = "Kit Schedule Move List" ' <-------- 1) change this if the sheet name of tyhe auto generated move list changes
    Dim located As Boolean, M_Located As Boolean
    
    'identifies size of the current data
    delete = list.UsedRange.Rows.Count
    
    'deletes all data from teh list starting from row 2 until the end as identified above
    Range(list.Cells(2, 1), list.Cells(delete, 5)).Clear

restart: ' <----jump point if it needs to be restarted.
    
    'searches each open workbook in Excel applicaiton to identify 1 with the predetermined sheet name if this name needs to be updated see above, number 1.
    For Each xworkbook In Application.Workbooks
        If xworkbook.Name <> home.Name Then ' ignores home workbook
            For Each sht In xworkbook.Worksheets 'steps through each sheet in the seleceted workbook.
                If sht.Name = SheetName Or sht.Name = "Instruments Schedule Move List" Then ' compares teh name of the sheet against the predetemined name 1
                    If sht.Cells(1, 2) Like ("Instruments Schedule Move List*") Then Set Malosa = xworkbook: M_Located = True: Exit For
                    If sht.Cells(1, 2) Like ("Kit Schedule Move List*") Then Set BVI = xworkbook: located = True: Exit For 'if correct sheet name checks the cell and checs for the text after the like, this will be the first part of the cell, it doesnot matter what follows.
                End If
            Next sht
            If located = True And M_Located = True Then Exit For ' if found steps out of changing workbook
        End If
    Next xworkbook
    
    'errors if can't find the sheet required
    If located = False Then
        MsgBox ("Ensure that you have downloaded and opened the automatic kit schedule move list, and enabled the content.") ' If it hasn't found the required sheet, throws an error.
        yesno = MsgBox("Do you want to try again?", vbYesNo) 'asks if the user wants to try again.
        If yesno = 6 Then GoTo restart Else Exit Sub ' vbYesNo gives a numerical outcome, 6 is yes, anyother answer exits the routine
    End If
    If M_Located = False Then
        MsgBox ("Ensure that you have downloaded and opened the Instruments Schedule Move List, and enabled the content.") ' If it hasn't found the required sheet, throws an error.
        yesno = MsgBox("Do you want to try again?", vbYesNo) 'asks if the user wants to try again.
        If yesno = 6 Then GoTo restart Else Exit Sub ' vbYesNo gives a numerical outcome, 6 is yes, anyother answer exits the routine
    End If
    
    'Copies the sheet into the working file, then closes the downloaded workbook
    BVI.Worksheets(SheetName).Copy Before:=home.Worksheets(1)
    BVI.Close SaveChanges:=False
    Set required = Worksheets(1) ' defines the sheet copied in as required for later calculations
    
    Malosa.Worksheets(1).Copy Before:=home.Worksheets(1)
    Malosa.Close SaveChanges:=False
    Set M_Required = Worksheets(1) ' defines the sheet copied in as required for later calculations
    
    outrow = 2 ' Starts the loop on row 2 
    
    Call Run(4, outrow, required, list, box_qty, 4, 7, 8)
    
    outrow = list.Cells(Rows.Count, 1).End(xlUp).Row + 1 ' Selects the next blank line
    
    Call Run(4, outrow, M_Required, list, box_qty, 4, 7, 8)
    
    ' consumables check
    boxes = box_qty.UsedRange.Rows.Count
    For i = 2 To Cons.UsedRange.Rows.Count
        If Cons.Cells(i, 3) > 0 Or Cons.Cells(i, 4) > 0 Then
            For j = 2 To boxes ' For each identified part cycles though all of the parts and box sizes present until it finds the part number
                If Trim(Cons.Cells(i, 1)) = Trim(box_qty.Cells(j, 1)) Then ' the tirm section removes any additional spaces at hte start / end and ensures everything is a strin whilst comparing, rather than trying to compare a string to number, it doesn't like that
                    list.Cells(outrow, 1) = Cons.Cells(i, 1) ' prints the required part number, i-2 because headings are only 1 line rather than 4.
                    If Cons.Cells(i, 4) > 0 Then 'if the pallet quantity is not 0 and there are pallets requested
                        If box_qty.Cells(j, 3) <> 0 Then 'if the pallet quantity is entered
                            list.Cells(outrow, 5) = Cons.Cells(i, 4) * box_qty.Cells(j, 3) 'times pallets by pallet quantity
                        Else: list.Cells(outrow, 5) = "Pallet Qty Needed" 'if it should be pallets but no pallet quantity states pallet qty needed
                        End If
                    ElseIf Cons.Cells(i, 3) > 0 Then
                        If box_qty.Cells(j, 2) <> 0 Then
                            list.Cells(outrow, 5) = Cons.Cells(i, 3) * box_qty.Cells(j, 2)
                        Else: list.Cells(outrow, 5) = "Box Qty Needed"
                        End If
                    Else: list.Cells(outrow, 5) = "Box Qty needed" ' states box ty needed if the box qty is blank or = 0
                    End If
                    outrow = outrow + 1
                    Exit For
                End If
            Next j
            If InStr(Cons.Cells(i, 2), "pallet") > 0 Then 'if there is no part number and the name includes pallets
                list.Cells(outrow, 1) = Cons.Cells(i, 2) ' copies pallet name
                list.Cells(outrow, 5) = Cons.Cells(i, 4) ' copies the number of pallets needed
                outrow = outrow + 1
            ElseIf list.Cells(outrow - 1, 1) <> Cons.Cells(i, 1) Then list.Cells(outrow, 1) = Cons.Cells(i, 1): list.Cells(outrow, 5) = "Box & Pallet Qty needed": outrow = outrow + 1 'if there are no box or pallet quantities for the part number then box & pallet quantity to be entered
            End If
        End If
    Next i

    Range(Cons.Cells(2, 3), Cons.Cells(Cons.UsedRange.Rows.Count, 4)).Clear ' clears the boxes & pallets populated
    
    'makes the column show as a date
    list.Columns(3).NumberFormat = "dd/mm/yyyy"
    ' stops alerts then deletes the requirements
    Application.DisplayAlerts = False
    required.delete
    M_Required.delete
    Application.DisplayAlerts = True ' starts the alerts again to stop me doing something stupid
    
    
    
End Sub