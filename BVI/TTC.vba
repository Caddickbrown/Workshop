Sub TTC()
        
    Sheets("Sheet1").Columns("B:B").TextToColumns Destination:=Range( _
        "Sheet1[[#Headers],[Part No]]"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True

End Sub
