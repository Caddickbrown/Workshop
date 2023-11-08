Sub ArchiveCompleted()
    Dim wsBVI As Worksheet
    Dim wsMalosa As Worksheet
    Dim wsComplete As Worksheet
    Dim tblBVI As ListObject
    Dim tblMalosa As ListObject
    Dim LastRow As Long
    Dim i As Long
    Dim Password As String
    
    ' Set the password for protecting and unprotecting sheets
    Password = "baconbutty"
    
    ' Define the destination worksheet as "Complete"
    Set wsComplete = ThisWorkbook.Sheets("Complete") ' Change "Complete" to the name of your destination sheet
    
    ' Unprotect the destination sheet
    wsComplete.Unprotect Password:=Password
    
    ' Set the source worksheets based on the provided names
    On Error Resume Next
    Set wsBVI = ThisWorkbook.Sheets("BVI Main")
    Set wsMalosa = ThisWorkbook.Sheets("Malosa Main")
    On Error GoTo 0
    
    If wsBVI Is Nothing Or wsMalosa Is Nothing Then
        MsgBox "One or both of the source sheets does not exist."
        Exit Sub
    End If
    
    ' Set the source tables based on the provided names
    On Error Resume Next
    Set tblBVI = wsBVI.ListObjects("Table2")
    Set tblMalosa = wsMalosa.ListObjects("Table6")
    On Error GoTo 0
    
    If tblBVI Is Nothing Or tblMalosa Is Nothing Then
        MsgBox "One or both of the source tables does not exist."
        Exit Sub
    End If
    
    ' Unprotect the source sheets
    wsBVI.Unprotect Password:=Password
    wsMalosa.Unprotect Password:=Password
    
    ' Find the last row in the source tables and move completed orders
    For Each tbl In Array(tblBVI, tblMalosa)
        For i = tbl.ListRows.Count To 1 Step -1
            If tbl.ListRows(i).Range.Cells(1, tbl.ListColumns("Status").Index).Value = "Completed" Then
                ' Copy the entire row to the destination sheet
                tbl.ListRows(i).Range.Copy wsComplete.Cells(wsComplete.Cells(wsComplete.Rows.Count, "A").End(xlUp).Row + 1, 1)
                
                ' Delete the row from the source table (optional)
                tbl.ListRows(i).Delete
            End If
        Next i
    Next tbl
    
    ' Protect the source sheets and the destination sheet again
    wsBVI.Protect Password:=Password, AllowSorting:=True, AllowFiltering:=True
    wsMalosa.Protect Password:=Password, AllowSorting:=True, AllowFiltering:=True
    wsComplete.Protect Password:=Password, AllowSorting:=True, AllowFiltering:=True
End Sub

