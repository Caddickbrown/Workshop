Sub FinanceReview()
    Cells.ClearContents
    Range("A1:M1").Value = Array("Notes", "AMZ", "S", "J", "M", "Notes", "", "", "AMZ", "S", "J", "M", "Total")
    Range("I2:M2").FormulaR1C1 = Array("=SUM(C[-7])", "=SUM(C[-7])", "=SUM(C[-7])", "=SUM(C[-7])", "=SUM(RC[-4]:RC[-1])")
    Range("A2").Select
End Sub

