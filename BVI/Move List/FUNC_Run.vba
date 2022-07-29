Function Run(start As Integer, outrow As Long, required As Worksheet, list As Worksheet, box_qty As Worksheet, expdate As Integer, Loc As Integer, pick As Integer)
    Dim Need As Integer, boxes As Integer
    'finds how many items we need to cycle through
    Need = required.UsedRange.Rows.Count
    boxes = box_qty.UsedRange.Rows.Count

    'cycles through the sheet of requirements to the end copies across teh required data
    For i = start To Need ' starts at row 4, below the headings
        If Trim(required.Cells(i, 1)) = "AMCO" Then ' makes sure that only the data from AMCO locations is run, if the location is updated in teh future update the text it is searching for.
            For j = 2 To boxes ' For each identified part cycles though all of the parts and box sizes present until it finds the part number
                If Trim(required.Cells(i, 3)) = Trim(box_qty.Cells(j, 1)) Then ' the tirm section removes any additional spaces at hte start / end and ensures everything is a strin whilst comparing, rather than trying to compare a string to number, it doesn't like that
                    list.Cells(outrow, 1) = required.Cells(i, 3) ' prints the required part number, i-2 because headings are only 1 line rather than 4.
                    list.Cells(outrow, 2) = required.Cells(i, 2) ' prints the required batch
                    list.Cells(outrow, 3) = required.Cells(i, expdate) 'prints teh expiry date
                    list.Cells(outrow, 4) = required.Cells(i, pick) ' prints th equantity that is required
                    If box_qty.Cells(j, 3) <> 0 And box_qty.Cells(j, 4) = "y" Then 'if the pallet quantity is not 0 and the part should be coming in as pallets
                        ' If the qty needed is greater than the amount in teh location takes everything from the location
                        If required.Cells(i, pick) >= required.Cells(i, Loc) Then
                            list.Cells(outrow, 5) = required.Cells(i, Loc)
                        ' if the required parts are in a round number for the pallet quantity lists the required numbers.
                        ElseIf required.Cells(i, pick) Mod box_qty.Cells(j, 3) = 0 Then list.Cells(outrow, 5) = required.Cells(i, pick)
                        'checks if the pallet quantities required are greater than the quantity in location if so gives locaiton quantity
                        ElseIf WorksheetFunction.RoundUp(required.Cells(i, pick) / box_qty.Cells(j, 3), 0) * box_qty.Cells(j, 3) > required.Cells(i, Loc) Then list.Cells(outrow, 5) = required.Cells(i, Loc)
                        'if pallet qty * number of pallets is possible states that volume is needed
                        Else: list.Cells(outrow, 5) = WorksheetFunction.RoundUp(required.Cells(i, pick) / box_qty.Cells(j, 3), 0) * box_qty.Cells(j, 3)
                        End If
                    ElseIf box_qty.Cells(j, 2) <> 0 Then ' confirms that the box size has been inputted and is greater than 0, if not it states box qty needed.
                        ' If the qty needed is greater than the amount in teh location takes everything from the location
                        If required.Cells(i, pick) >= required.Cells(i, Loc) Then
                            list.Cells(outrow, 5) = required.Cells(i, Loc)
                        ' if the required parts are in a round number for the box quantity lists the required numbers.
                        ElseIf required.Cells(i, pick) Mod box_qty.Cells(j, 2) = 0 Then list.Cells(outrow, 5) = required.Cells(i, pick)
                        'checks if the box quantities required are greater than the quantity in location if so gives locaiton quantity
                        ElseIf WorksheetFunction.RoundUp(required.Cells(i, pick) / box_qty.Cells(j, 2), 0) * box_qty.Cells(j, 2) > required.Cells(i, Loc) Then list.Cells(outrow, 5) = required.Cells(i, Loc)
                        'if box qty * number of boxes is possible states that volume is needed
                        Else: list.Cells(outrow, 5) = WorksheetFunction.RoundUp(required.Cells(i, pick) / box_qty.Cells(j, 2), 0) * box_qty.Cells(j, 2)
                        End If
                    Else: list.Cells(outrow, 5) = "Box Qty needed" ' states box ty needed if the box qty is blank or = 0
                    End If
                    outrow = outrow + 1
                    Exit For
                End If
            Next j
            'if the part number does not appear in teh box quantity publishes the parts needed and quantities and states box qty needed
            If list.Cells(outrow - 1, 1) <> required.Cells(i, 3) And list.Cells(outrow - 1, 2) <> required.Cells(i, 2) Then list.Cells(outrow, 1) = required.Cells(i, 3): list.Cells(outrow, 2) = required.Cells(i, 2): list.Cells(outrow, 3) = required.Cells(i, expdate): list.Cells(outrow, 4) = required.Cells(i, pick): list.Cells(outrow, 5) = "Box Qty needed": outrow = outrow + 1
        End If
    Next i

End Function
