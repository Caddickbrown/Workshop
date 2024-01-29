Sub partsRemovalList()

'Dim t As Single
't = Timer

    Dim PartsRem As Worksheet, PartsPivot As Worksheet, AllParts As Worksheet, land As Worksheet
    Set PartsRem = Worksheets("Parts to be removed")
    Set PartsPivot = Worksheets("PartsPivot")
    Set AllParts = Worksheets("AllParts")
    Set land = Worksheets("Coversheet")
    
    Range(PartsRem.Cells(2, 1), PartsRem.Cells(PartsRem.UsedRange.Rows.Count, 3)).Clear ' clears the parts to be removed sheet
        
    
    Dim remnum As Integer, located As Boolean, outstep As Long
    remnum = land.Cells(2, 2) 'sets the maximum number of intactions planed before they can be removed
    outstep = 2
    
    For Part = 2 To PartsPivot.UsedRange.Rows.Count ' steps through all teh parts from teh warehouse
        For step = 3 To AllParts.UsedRange.Rows.Count ' checks if each part is needed
            If AllParts.Cells(step, 1) = PartsPivot.Cells(Part, 1) Then
                located = True
                If AllParts.Cells(step, 2) <= remnum Then ' if its needed checks if the interactions planned are above the rset number
                    PartsRem.Cells(outstep, 1) = PartsPivot.Cells(Part, 1) 'if not copies the part No, qty in wh & the number of interactions
                    PartsRem.Cells(outstep, 2) = PartsPivot.Cells(Part, 2)
                    PartsRem.Cells(outstep, 3) = AllParts.Cells(step, 2)
                    outstep = outstep + 1
                End If
                Exit For
            End If
        Next step
        If located = True Then
            located = False
        Else
            PartsRem.Cells(outstep, 1) = PartsPivot.Cells(Part, 1) 'if the part is not needed copies the part No, qty in wh & 0 interactions to removals list
            PartsRem.Cells(outstep, 2) = PartsPivot.Cells(Part, 2)
            PartsRem.Cells(outstep, 3) = 0
            outstep = outstep + 1
        End If
    Next Part

    PartsRem.Select

'MsgBox Format((Timer - t) / 86400, "hh:mm:ss")

End Sub