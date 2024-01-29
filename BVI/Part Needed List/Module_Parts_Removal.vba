'OPen IFS shop Orders
'Search Site: 2051 & Shop Order Status: Released;Planned;Started & Part No: 581%;582%;583%;584%;585%;8%;%BID;MMSU% on IFS.
'Download data
'open IFS search Manufacturing structures
'Search SIte: 2051 & Parent Part No:588%;589%;59%;NS%;MMK% & Parent part Status: A & Status: Buildable  on IFS.
'Download Data
'open IFS Inventory Part in stock
' search site:2051 & Warehouse: GOODS-IN;2 & Bay: !=%AMCO%
'download data

Sub partsRemoval()

'Dim t As Single
't = Timer

'all relevant sheets needed
    Dim home As Workbook, AllParts As Worksheet, Need As Worksheet, InsNeed As Worksheet, Kanbans As Worksheet, land As Worksheet, PartsRem As Worksheet, dataout As Worksheet, PT As PivotTable, PartsPivot As Worksheet
    Set home = ThisWorkbook
    Set Need = Worksheets("Parts Needed")
    Set Kanbans = Worksheets("Kanbans")
    Set land = Worksheets("Coversheet")
    Set PartsPivot = Worksheets("PartsPivot")
    Set InsNeed = Worksheets("InsPartNeed")
    Set AllParts = Worksheets("AllParts")
    
    Call LocateSheet("OverviewInventoryPartInStock", "Location No", "Warehouse", dataout, home, "2", "GOODS-IN", "BLANK") ' locates teh parts in stock sheet
    
    Call CreatePIvot(dataout, home, PartsPivot, PT) 'creates the parts pivot
    
    'populate pivot table with required data
    With PT
        .PivotFields("Part No").Orientation = xlRowField
        With .PivotFields("On Hand Qty")
            .Orientation = xlDataField
            .Position = 1
            .Function = xlSum
        End With
    End With
    
    'Deletes IFS parts in stock list
    Application.DisplayAlerts = False 'turns off pop-ups to ask if ok to delete sheet
        dataout.Delete 'Deletes sheet
    Application.DisplayAlerts = True 'turns pop-ups back on to stop me doing something silly

    ' pulls out the required parts needed from teh shop orders and BOMs into worksheets as specified below
    Call PartNeed(home, Worksheets("InsPivotOut"), Worksheets("InsExtract"), Worksheets("InsBom"), InsNeed) ' findst eh number of parts needed to generate make the released instruments

    Set dataout = home.Worksheets.Add(Worksheets(1)) ' creates a new sheet
    'creates titles for later pivot
    dataout.Cells(1, 1) = "Part No"
    dataout.Cells(1, 2) = "Interactions"
    dataout.Cells(1, 3) = "Qty"
    
    'sets output row location
    Dim outstep As Long
    outstep = 2
    
    ' copiest eh required parts needed to make kits into temp dataout sheet
    For step = 3 To Need.UsedRange.Rows.Count - 1
        For j = 1 To 3
            dataout.Cells(outstep, j) = Need.Cells(step, j)
        Next j
        outstep = outstep + 1
    Next step
    ' copies the required parts needed to make the instruments into temp dataout sheet
    For step = 3 To InsNeed.UsedRange.Rows.Count - 1
        For j = 1 To 3
            dataout.Cells(outstep, j) = InsNeed.Cells(step, j)
        Next j
        outstep = outstep + 1
    Next step
    
    Call CreatePIvot(dataout, home, AllParts, PT) 'creates the all parts needed for kits & instruments pivot
    
    'populate pivot table with required data
    With PT
        .PivotFields("Part No").Orientation = xlRowField
        With .PivotFields("Interactions")
            .Orientation = xlDataField
            .Position = 1
            .Function = xlSum
            
        End With
        With .PivotFields("Qty")
            .Orientation = xlDataField
            .Position = 2
            .Function = xlSum
        End With
    End With
    
    'Deletes IFS parts in stock list
    Application.DisplayAlerts = False 'turns off pop-ups to ask if ok to delete sheet
        dataout.Delete 'Deletes sheet
    Application.DisplayAlerts = True 'turns pop-ups back on to stop me doing something silly
    
    
    Call partsRemovalList
       
'MsgBox Format((Timer - t) / 86400, "hh:mm:ss")
    
End Sub
