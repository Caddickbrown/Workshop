Sub DeleteRows()
' This deletes rows where the Component Requirements are 0 or lower (Column D is for Component Requirement) - basically, where we have enough material

    Dim i As Long
    For i = Range("D" & Rows.Count).End(xlUp).Row To 1 Step -1
        If Not (Range("D" & i).Value > 0) Then
            Range("D" & i).EntireRow.Delete
        End If
    Next i

End Sub


Sub DeleteRows()
    'get last row in column A
    Last = Cells(Rows.Count, "D").End(xlUp).Row
    For i = Last To 1 Step -1
        'if cell value is less than 100
        If (Cells(i, "D").Value) < 100 Then
            'delete entire row
            Cells(i, "D").EntireRow.Delete
        End If
    Next i
End Sub


With Sheets(1).UsedRange
    For lrow = .Rows.Count To 2 Step -1
        If .Cells(lrow, 4).Value <> 110 Then .Rows(lrow).Delete
    Next lrow
End With