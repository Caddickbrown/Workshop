'ToDo
'- [ ] Add calling Subs/Abstractions
'- [ ] Looping?
Option Explicit

Private SheetPassword As String
Set SheetPassword = "baconbutty"
Private sheettarget As String
Private tabletarget As String

Sub UnlockSheet(UnlockTarget as String)

    Worksheets(UnlockTarget).Unprotect Password:=SheetPassword

End Sub

Sub LockSheet(LockTarget as String)

    Worksheets(LockTarget).Protect Password:=SheetPassword, AllowFiltering:=True

End Sub

Sub RevealAllRows()
    
    'Unhide/Clear Filters
    Rows("1:1048576").EntireRow.Hidden = False
    If ActiveSheet.FilterMode = True Then
        ActiveSheet.ShowAllData
    End If

End Sub

Sub SortingHat(UnlockTargetTab As String, TargetTable As String)

    Sequence Sort UnlockTargetTab:=UnlockTargetTab, TargetTable:=TargetTable

End Sub





'New Schedule Sort
Sub BVISortSheet()

    Set sheettarget = "BVI Main"
    Set tabletarget = "Table2"

    UnlockSheet sheettarget
    ScheduleSort
    LockSheet

End Sub



Sub ScheduleSort()
'Sorts the Sheet into Schedule Order

    Worksheets(sheettarget).Unprotect Password:="baconbutty" 'Unprotect the Sheet with the password

    'Unhide/Clear Filters
    Rows("1:1048576").EntireRow.Hidden = False
    If ActiveSheet.FilterMode = True Then
        ActiveSheet.ShowAllData
    End If
    
    ActiveWorkbook.Worksheets(sheettarget).ListObjects(tabletarget).Sort.SortFields.Add2 _
        Key:=Range(tabletarget & "[[#All],[Sequence]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(sheettarget).ListObjects(tabletarget).Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ActiveWorkbook.Worksheets(sheettarget).ListObjects(tabletarget).Sort.SortFields.Clear
    
    ActiveWorkbook.Worksheets(sheettarget).ListObjects(tabletarget).Sort.SortFields.Add2 _
        Key:=Range(tabletarget & "[[#All],[Date]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(sheettarget).ListObjects(tabletarget).Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub

Sub MalosaScheduleSort()

    Worksheets("Malosa Main").Unprotect Password:="baconbutty" 'Unprotect the Sheet with the password

    Rows("1:1048576").EntireRow.Hidden = False 'Unhide any rows

    'Clear Filters
    If ActiveSheet.FilterMode = True Then
        ActiveSheet.ShowAllData
    End If


    'Sort on Sequence
    ActiveWorkbook.Worksheets("Malosa Main").ListObjects("Table6").Sort.SortFields.Add2 _
        Key:=Range("Table6[[#All],[Sequence]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Malosa Main").ListObjects("Table6").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'Sort on Date
    ActiveWorkbook.Worksheets("Malosa Main").ListObjects("Table6").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Malosa Main").ListObjects("Table6").Sort.SortFields.Add2 _
        Key:=Range("Table6[[#All],[Date]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Malosa Main").ListObjects("Table6").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ActiveSheet.Protect Password:="baconbutty", AllowFiltering:=True 'Protect the sheet with the password, allowing filtering

End Sub