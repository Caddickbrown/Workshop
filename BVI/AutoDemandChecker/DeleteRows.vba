Sub DeleteRows()
' This deletes rows where the Component Requirements are 0 or lower (Column D is for Component Requirement) - basically, where we have enough material

    Dim i As Long
    For i = Range("D" & Rows.Count).End(xlUp).Row To 1 Step -1
        If Not (Range("D" & i).Value > 0) Then
            Range("D" & i).EntireRow.Delete
        End If
    Next i

End Sub