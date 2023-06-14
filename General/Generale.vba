' Check if your sheet is read only and prompt if it is
    If Home.ReadOnly Then
        i = MsgBox("SIC sheet is opened as Read Only. Please reopen the file and ensure that it is not Read Only before trying to run. Do you wish to continue?", vbYesNo)
        If i = 6 Then MsgBox ("Please save file under a different name."): Exit Sub
        If i = 7 Then Home.Close
    End If








