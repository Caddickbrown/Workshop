'OPen IFS shop Orders
'Search Site: 2051 & Shop Order Status: Released;Planned;Started & Part No: 588%;589%;59%;NS%;MMK% on IFS.
'Download data
'open IFS search Manufacturing structures
'Search SIte: 2051 & Parent Part No:588%;589%;59%;NS%;MMK% & Parent part Status: A & Status: Buildable  on IFS.
'Download Data

Sub Part_Rev_List()
    
Dim t As Single
t = Timer
        
    Dim home As Workbook, dataout As Worksheet, extract As Worksheet, out As Worksheet, BOM As Worksheet, Need As Worksheet, PT As PivotTable, PivotOut As Worksheet, Kanbans As Worksheet, land As Worksheet
    Set home = ThisWorkbook
    Set extract = Worksheets("Extracted")
    Set BOM = Worksheets("Needed BOMs")
    Set Need = Worksheets("Parts Needed")
    Set PivotOut = Worksheets("PivotOut")
    Set Kanbans = Worksheets("Kanbans")
    Set land = Worksheets("Coversheet")
    
    'Application.StatusBar = "Importing data set 1 of 2 0%"
    
    Call LocateSheet("OverviewShopOrder", "Part No", "Shop Order Status", dataout, home, "Started", "Released", "Planned")
         
    'Application.StatusBar = "Importing data set 1 of 2 50%"
         
    Call CreatePIvot(dataout, home, PivotOut, PT) 'creates the needed pivot
    
    'populate pivot table with required data
    With PT
        .PivotFields("Part No").Orientation = xlRowField
        .PivotFields("Part Revision").Orientation = xlRowField
        With .PivotFields("Part Description")
            .Orientation = xlDataField
            .Position = 1
            .Function = xlCount
        End With
        With .PivotFields("Lot Size")
            .Orientation = xlDataField
            .Position = 2
            .Function = xlSum
        End With
    End With
      
    'Application.StatusBar = "Importing data set 1 of 2 90%"
      
    'Deletes IFS shop order list
    Application.DisplayAlerts = False 'turns off pop-ups to ask if ok to delete sheet
        dataout.Delete 'Deletes sheet
    Application.DisplayAlerts = True 'turns pop-ups back on to stop me doing something silly
      
    Application.StatusBar = "Importing data set 1 of 2 100%"
    
    Dim length As Long, loc As Long, outPos As Long
    
    length = PivotOut.UsedRange.Rows.Count
    outPos = 2
     ' clears the list before it starts to run again
    Range(extract.Cells(2, 1), extract.Cells(extract.UsedRange.Rows.Count, 5)).Clear
    'steps through all the kits needed and the revision and produces 1 list including the number of iterations and total qty needed.
    For i = 3 To length
        If PivotOut.Cells(i, 1) > 1000 Then 'etither if teh number is above 1000 or the part number is nonsterile then set as location for the pack
            loc = i
        ElseIf Left(PivotOut.Cells(i, 1), 2) = "NS" Then ' i
            loc = i
        ElseIf PivotOut.Cells(i, 1) < 100 Then
            extract.Cells(outPos, 1) = PivotOut.Cells(loc, 1)
            For j = 1 To 3
                extract.Cells(outPos, 1 + j) = PivotOut.Cells(i, j)
            Next j
            extract.Cells(outPos, 5) = extract.Cells(outPos, 1) & "-" & extract.Cells(outPos, 2)
            outPos = outPos + 1
        End If
        
    'Application.StatusBar = "Creating kit list " & Format(i / length, "0.0%")
        
    Next i
     
    
    Call LocateSheet("OverviewManufacturingStructure", "Parent Part No", "Status", dataout, home, "Buildable", "Obsolete", "Cancelled")
    
    Application.StatusBar = "Importing data set 2 of 2 100%"
    
    ' clears the list before it starts to run again
    Range(BOM.Cells(2, 1), BOM.Cells(BOM.UsedRange.Rows.Count, 5)).Clear
    'identifies the required parts needed from the list of BOMs
    length = dataout.UsedRange.Rows.Count
    ColumnCount = dataout.Cells(1, Columns.Count).End(xlToLeft).Column
    
    
    For i = 1 To ColumnCount
        If dataout.Cells(1, i) = "Parent Part No" Then
            PP = i
            If PP > 0 And Rev > 0 And CP > 0 And BQ > 0 Then Exit For
        ElseIf dataout.Cells(1, i) = "Revision" Then
            Rev = i
            If PP > 0 And Rev > 0 And CP > 0 And BQ > 0 Then Exit For
        ElseIf dataout.Cells(1, i) = "Component Part" Then
            CP = i
            If PP > 0 And Rev > 0 And CP > 0 And BQ > 0 Then Exit For
        ElseIf dataout.Cells(1, i) = "Qty per Assembly" Then
            BQ = i:
            If PP > 0 And Rev > 0 And CP > 0 And BQ > 0 Then Exit For
        End If
    Next i
    
    
    Dim located As Boolean, OutRow As Long, interactions As Integer, qty As Long
    located = False
    OutRow = 2
    For i = 2 To outPos
        interactions = extract.Cells(i, 3)
        qty = extract.Cells(i, 4)
        For j = 2 To length
            If extract.Cells(i, 5) = dataout.Cells(j, PP) & "-" & dataout.Cells(j, Rev) Then 'finds the parts needed for the specified parts and revision fromt eh Manufacturing structures
                located = True ' identifies that the required pack has been located
                BOM.Cells(OutRow, 1) = dataout.Cells(j, PP) ' copies pack number
                BOM.Cells(OutRow, 2) = dataout.Cells(j, Rev) ' copies revision
                BOM.Cells(OutRow, 3) = dataout.Cells(j, CP)  'copies component part number
                BOM.Cells(OutRow, 4) = dataout.Cells(j, BQ) * qty ' copies qty needed per assembly
                BOM.Cells(OutRow, 5) = interactions
                OutRow = OutRow + 1
            ElseIf located = True Then located = False: Exit For
            End If
        Next j
        'Application.StatusBar = "Extracting Parts " & Format(i / outPos, "0.0%")
        'BOM.Cells(1, 10) = i / outPos
    Next i
    
    'Deletes IFS Manufacturing structures
    Application.DisplayAlerts = False 'turns off pop-ups to ask if ok to delete sheet
        dataout.Delete 'Deletes sheet
    Application.DisplayAlerts = True 'turns pop-ups back on to stop me doing something silly
    
    'Application.StatusBar = "Processing Parts 0%"
    
    Call CreatePIvot(BOM, home, Need, PT)
    
    'Application.StatusBar = "Processing Parts 50%"
    
    'populate pivot table with required data
    With PT
        .PivotFields("Component").Orientation = xlRowField
        With .PivotFields("Interactions")
            .Orientation = xlDataField
            .Position = 1
            .Function = xlSum
            
        End With
        With .PivotFields("Qty Needed")
            .Orientation = xlDataField
            .Position = 2
            .Function = xlSum
        End With
        .PivotFields("Component").AutoSort Order:=xlDescending, Field:="Sum of Interactions"
    End With
    
    'Application.StatusBar = "Processing Parts 100%"
    
    Call Parts_List
    
MsgBox Format((Timer - t) / 86400, "hh:mm:ss")
    
End Sub
