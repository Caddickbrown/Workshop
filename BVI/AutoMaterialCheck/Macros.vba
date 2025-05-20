Sub ProcessBlankCellsInMain()
    Dim iter1 As Long, iter2 As Long, iterations As Long
    Dim mainSheet As Worksheet
    Dim stockTallySheet As Worksheet
    Dim demandSheet As Worksheet
    Dim i As Long, j As Long
    Dim lastRow As Long, lastRowDemand As Long
    Dim valueToSearch As Variant
    
    ' Set references to worksheets
    Set mainSheet = ThisWorkbook.Worksheets("Main")
    Set stockTallySheet = ThisWorkbook.Worksheets("StockTally")
    Set demandSheet = ThisWorkbook.Worksheets("Demand")
    
    ' Get iterations value from cell X1 in StockTally sheet
    iter1 = mainSheet.Range("M2").Value
    iter2 = mainSheet.Range("O14").Value
    iterations = Application.WorksheetFunction.Min(iter1, iter2)
    mainSheet.Range("M2").Value = iterations
    mainSheet.Range("M3").Value = iterations
    
    mainSheet.Range("M5").Formula = "=NOW()"
    mainSheet.Range("M5").Formula = mainSheet.Range("M5").Value
    mainSheet.Range("M6").Formula = "=NOW()"
    
    ' Find the last row in column E of Main sheet
    lastRow = mainSheet.Cells(mainSheet.Rows.Count, "H").End(xlUp).Row
    If lastRow < mainSheet.Cells(mainSheet.Rows.Count, "A").End(xlUp).Row Then
        lastRow = mainSheet.Cells(mainSheet.Rows.Count, "A").End(xlUp).Row
    End If
    
    ' Find the last row in Demand sheet
    lastRowDemand = demandSheet.Cells(demandSheet.Rows.Count, "A").End(xlUp).Row
    
    Dim demandDict As Object
    Set demandDict = CreateObject("Scripting.Dictionary")
    
    Dim key As Variant
    For j = 1 To lastRowDemand
        key = demandSheet.Cells(j, "A").Value
        If Len(Trim(key)) > 0 Then
            If Not demandDict.exists(key) Then
                demandDict.Add key, j
            End If
        End If
    Next j
    
    ' Loop through column E in Main sheet
    For i = 1 To lastRow
        ' Check if the cell in column E is blank and we still have iterations left
        If IsEmpty(mainSheet.Cells(i, "H").Value) And iterations > 0 Then
            ' Copy values from columns A, B, D to StockTally cells F2, G2, H2
            stockTallySheet.Range("G2").Value = mainSheet.Cells(i, "A").Value
            stockTallySheet.Range("H2").Value = mainSheet.Cells(i, "B").Value
            stockTallySheet.Range("I2").Value = mainSheet.Cells(i, "F").Value
            
            ' Allow Excel to calculate
            Application.Calculate
            
            ' Wait for the Calculation for finish
            Do While Application.CalculationState <> xlDone
                DoEvents
            Loop
            
            ' Wait 2 seconds (May no longer be needed)
            ' Application.Wait Now + TimeValue("00:00:02")
            
            ' Copy value from T2 in StockTally to the blank cell in column E
            mainSheet.Cells(i, "H").Value = stockTallySheet.Range("T2").Value
            mainSheet.Cells(i, "I").Value = stockTallySheet.Range("T4").Value
            
            ' Check if T2 says "Release"
            If stockTallySheet.Range("T2").Value = "Release" Then
                valueToSearch = stockTallySheet.Range("G2").Value
                If demandDict.exists(valueToSearch) Then
                    demandSheet.Cells(demandDict(valueToSearch), "E").Value = "Released"
                End If
            End If
            
            ' Decrement iterations counter
            iterations = iterations - 1
            
            ' Update the iterations value in StockTally sheet
            mainSheet.Range("M3").Value = iterations
            
            ' Check if we've done all iterations
            If iterations <= 0 Then
                Exit For
            End If
        End If
    Next i
    
    
    mainSheet.Range("M6").Formula = mainSheet.Range("M6").Value

End Sub


Sub TTC()
    
    Sheets("Demand").Columns("A:A").TextToColumns Destination:=Range("Demand[[#Headers],[SO No]]"), _
        DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter _
        :=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, _
        Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    
    Sheets("Demand").Columns("B:B").TextToColumns Destination:=Range("Demand[[#Headers],[Part No]]"), _
        DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter _
        :=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, _
        Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    
    Sheets("IPIS").Columns("A:A").TextToColumns Destination:=Range("IPIS[[#Headers],[PART_NO]]"), _
        DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter _
        :=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, _
        Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    
    Sheets("ManStructures").Columns("A:A").TextToColumns Destination:=Range( _
        "Manufacturing_Structures[[#Headers],[Parent Part]]"), DataType:=xlDelimited _
        , TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    
    Sheets("ManStructures").Columns("B:B").TextToColumns Destination:=Range( _
        "Manufacturing_Structures[[#Headers],[Component Part]]"), DataType:= _
        xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, _
        Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False _
        , FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    
    Sheets("Component Demand").Columns("B:B").TextToColumns Destination:=Range( _
        "Component_Demand[[#Headers],[Kit Number]]"), DataType:= _
        xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, _
        Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False _
        , FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    
    Sheets("Component Demand").Columns("C:C").TextToColumns Destination:=Range( _
        "Component_Demand[[#Headers],[Component Part Number]]"), DataType:= _
        xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, _
        Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False _
        , FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    
    Sheets("POs").Columns("B:B").TextToColumns Destination:=Range( _
        "POs[[#Headers],[Part Number]]"), DataType:= _
        xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, _
        Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False _
        , FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    
    ActiveWorkbook.Worksheets("Demand").ListObjects("Demand").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Demand").ListObjects("Demand").Sort.SortFields.Add2 _
        key:=Range("Demand[[#All],[Status]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Demand").ListObjects("Demand").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
                
    Sheets("Main").Select
End Sub


Sub FindAndReplacePatterns()
    ' This macro processes a list of values and handles different pattern replacements
    ' Each value in the list will be searched for in multiple formats (with/without semicolons)
    
    Dim ws As Worksheet
    Dim valueRange As Range
    Dim valueToProcess As String
    Dim i As Long
    Dim j As Long
    Dim searchRange As Range
    Dim foundCell As Range
    Dim firstAddress As String
    Dim findPatterns(1 To 4) As String
    Dim replacePatterns(1 To 4) As String
    
    ' Set the worksheets to work with
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    
    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets("Main")
    Set wsTarget = ThisWorkbook.Worksheets("Main")
    On Error GoTo 0
    
    ' Check if the specified sheets exist
    If wsSource Is Nothing Then
        MsgBox "The sheet 'Main' could not be found!", vbExclamation
        Exit Sub
    End If
    
    If wsTarget Is Nothing Then
        MsgBox "The sheet 'Main' could not be found!", vbExclamation
        Exit Sub
    End If
    
    ' Automatically determine the range containing the values to process
    ' This will find all consecutive non-empty cells in column N starting from N1 in "BOM Check" sheet
    Dim lastRow As Long
    lastRow = wsSource.Cells(wsSource.Rows.Count, "AB").End(xlUp).Row
    Set valueRange = wsSource.Range("AB1:AB" & lastRow)
    
    ' Define the range to search in (only column R in the "Main" sheet)
    Dim lastRowTarget As Long
    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "G").End(xlUp).Row
    Set searchRange = wsTarget.Range("H1:H" & lastRowTarget)
    
    ' Loop through each value in the list
    For i = 1 To valueRange.Rows.Count
        valueToProcess = Trim(valueRange.Cells(i, 1).Value)
        
        ' Skip empty values
        If Not IsEmpty(valueToProcess) Then
            ' Create the different patterns to find/replace for this value
            ' Pattern 1: ";value;" -> ";"
            findPatterns(1) = ";" & valueToProcess & ";"
            replacePatterns(1) = ";"
            
            ' Pattern 2: ";value" -> ""
            findPatterns(2) = ";" & valueToProcess
            replacePatterns(2) = ""
            
            ' Pattern 3: "value;" -> ""
            findPatterns(3) = valueToProcess & ";"
            replacePatterns(3) = ""
            
            ' Pattern 4: "value" -> ""
            findPatterns(4) = valueToProcess
            replacePatterns(4) = ""
            
            ' Process each pattern for this value
            For j = 1 To 4
                ' Initialize for a new search
                Set foundCell = Nothing
                
                ' Find all instances of this pattern and replace them
                Set foundCell = searchRange.Find(What:=findPatterns(j), LookIn:=xlValues, _
                                              LookAt:=xlPart, SearchOrder:=xlByRows, _
                                              SearchDirection:=xlNext, MatchCase:=True)
                
                ' If the pattern is found, replace it and continue searching
                If Not foundCell Is Nothing Then
                    firstAddress = foundCell.Address
                    
                    Do
                        ' Store the original value
                        Dim originalValue As String
                        originalValue = foundCell.Value
                        
                        ' Replace only the found pattern in the cell, not the entire cell value
                        foundCell.Value = Replace(originalValue, findPatterns(j), replacePatterns(j))
                        
                        ' Find the next occurrence
                        Set foundCell = searchRange.FindNext(After:=foundCell)
                        
                        ' Exit loop if no more matches or we've gone full circle
                        If foundCell Is Nothing Then Exit Do
                        If foundCell.Address = firstAddress Then Exit Do
                    Loop
                End If
            Next j
            
            Application.StatusBar = "Processed: " & valueToProcess
        End If
    Next i
    
    Application.StatusBar = False
    MsgBox "Find and replace operations completed!", vbInformation
End Sub



