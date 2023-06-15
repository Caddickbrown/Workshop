' ## To Do
' - [ ] Variable Tab Names
' - [ ] Auto Pull Data sheets
' - [ ] 
' - [ ] 
' - [ ] 
' - [ ] 
' - [ ] 



Sub MasterSheetRefresh()

Sheets("Locations").Rows("3:1048576").Delete Shift:=xlUp
Sheets("Requisition Demand").Rows("2:1048576").ClearContents
Sheets("Released Shop Orders").Rows("2:1048576").ClearContents
Sheets("IPIS").Cells.ClearContents

Sheets("Locations").Range("A1:N1").Value = Array("Part Number", "Total Raw Material Qty", "AMCO", "GOODS-IN", "INST&KNIVES", "CENTRAL-STORES", "B1 Stock", "RM Material", "Total Req For Week", "RM Shortage", "B1 Shortage", "Quick Release", "Released SO", "Net Usable RM")
Sheets("Locations").Range("A2:N2").Formula2 = Array("='Requisition Demand'!A2", "=SUMIF(INDEX(IPIS!$A:$BZ,0,MATCH(""Part No"",IPIS!$A$1:$BZ$1,0)),LEFT(A2,8)&""A"",INDEX(IPIS!$A:$BZ,0,MATCH(""On Hand Qty"",IPIS!$A$1:$BZ$1,0)))", "=SUMIFS(INDEX(IPIS!$A:$BZ,0,MATCH(""On Hand Qty"",IPIS!$A$1:$BZ$1,0)),INDEX(IPIS!$A:$BZ,0,MATCH(""Warehouse"",IPIS!$A$1:$BZ$1,0)),C$1,INDEX(IPIS!$A:$BZ,0,MATCH(""Part No"",IPIS!$A$1:$BZ$1,0)),LEFT($A2,8)&""A"")", "", "", "", "=E2+F2", "=CONCATENATE(LEFT(A2,8),""A"")", "=VLOOKUP(A2,'Requisition Demand'!A:B,2,0)", "=B2-I2", "=G2-I2", "=MIN(I2,G2-M2)", "=SUMIF('Released Shop Orders'!A:A,A2,'Released Shop Orders'!B:B)", "=G2-M2")
Sheets("Locations").Range("C2").AutoFill Destination:=Range(Cells(2, 3), Cells(2, 6))

Sheets("Requisition Demand").Range("A1:C1").Value = Array("Part Numbers", "Sum of Quantity", "Priority")

Sheets("Released Shop Orders").Range("A1:B1").Value = Array("Part Numbers", "Lot Size", "Priority")

Sheets("Locations").Range("A2").Select

End Sub
