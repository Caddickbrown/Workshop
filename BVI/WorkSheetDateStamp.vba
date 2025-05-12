Private Sub Worksheet_Change(ByVal Target As Excel.Range)
        With Target
            If .Count > 1 Then Exit Sub
            If Not Intersect(Range("B2:B328"), .Cells) Is Nothing Then

                    ActiveSheet.Unprotect "QualityBVI"
                    With .Offset(0, 11)
                        .NumberFormat = "dd mmm hh:mm"
                        .Value = Now
                    End With
                    With .Offset(0, 1)
                    .Select
                    End With
                    ActiveSheet.Protect "QualityBVI"
                End If

        End With
    End Sub



    Private Sub Worksheet_Change(ByVal Target As Range)
    Dim Cell As Range

    ' Exit if more than one cell is changed
    If Target.CountLarge > 1 Then Exit Sub

    ' Check if the changed cell intersects with B2:B328
    If Not Intersect(Target, Me.Range("B2:B328")) Is Nothing Then
        Application.EnableEvents = False ' Prevent recursive triggering

        ' Unprotect the sheet
        Me.Unprotect Password:="QualityBVI"

        ' Timestamp 11 columns to the right of the changed cell
        With Target.Offset(0, 11)
            .NumberFormat = "dd mmm hh:mm"
            .Value = Now
        End With

        ' Optional: Select one column to the right of the changed cell (might not be necessary)
        Target.Offset(0, 1).Select

        ' Reprotect the sheet
        Me.Protect Password:="QualityBVI"

        Application.EnableEvents = True
    End If
End Sub
