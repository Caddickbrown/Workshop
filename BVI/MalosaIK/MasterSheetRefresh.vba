' ## To Do
' - [ ] Auto Pull Exported Data sheets
' - [ ] Formatting Reset
' - [ ] Variable Formulas in Arrays
' - [ ] 
' - [ ] 
' - [ ] 

Sub MasterSheetRefresh()

Dim MainSheetName As String, RequisitionsSheetName As String, ReleasedOrdersSheetName As String, IPISSheetName As String

MainSheetName = "Locations"
RequisitionsSheetName = "Requisition Demand"
ReleasedOrdersSheetName = "Released Shop Orders"
IPISSheetName = "IPIS"

' ## Space for Formulas
' PartNumberCalc = 
' TotalRawMaterialQtyCalc
' AMCOCalc
' GoodsInCalc
' InstrumentsAndKnivesCalc
' CentralStoresCalc
' B1StockCalc
' RMMaterialCalc
' TotalReqForWeekCalc
' RMShortageCalc
' B1ShortageCalc
' QuickReleaseCalc
' ReleasedSOCalc
' NetUsableRMCalc

Sheets(MainSheetName).Rows("3:1048576").Delete Shift:=xlUp
Sheets(RequisitionsSheetName).Rows("2:1048576").ClearContents
Sheets(ReleasedOrdersSheetName).Rows("2:1048576").ClearContents
Sheets(IPISSheetName).Cells.ClearContents

Sheets(MainSheetName).Range("A1:N1").Value = Array("Part Number", "Total Raw Material Qty", "AMCO", "GOODS-IN", "INST&KNIVES", "CENTRAL-STORES", "B1 Stock", "RM Material", "Total Req For Week", "RM Shortage", "B1 Shortage", "Quick Release", "Released SO", "Net Usable RM")
Sheets(MainSheetName).Range("A2:N2").Formula2 = Array("='" & RequisitionsSheetName & "'!A2", "=SUMIF(INDEX(" & IPISSheetName & "!$A:$BZ,0,MATCH(""Part No""," & IPISSheetName & "!$A$1:$BZ$1,0)),LEFT(A2,8)&""A"",INDEX(" & IPISSheetName & "!$A:$BZ,0,MATCH(""On Hand Qty""," & IPISSheetName & "!$A$1:$BZ$1,0)))", "=SUMIFS(INDEX(" & IPISSheetName & "!$A:$BZ,0,MATCH(""On Hand Qty""," & IPISSheetName & "!$A$1:$BZ$1,0)),INDEX(" & IPISSheetName & "!$A:$BZ,0,MATCH(""Warehouse""," & IPISSheetName & "!$A$1:$BZ$1,0)),C$1,INDEX(" & IPISSheetName & "!$A:$BZ,0,MATCH(""Part No""," & IPISSheetName & "!$A$1:$BZ$1,0)),LEFT($A2,8)&""A"")", "", "", "", "=E2+F2", "=CONCATENATE(LEFT(A2,8),""A"")", "=VLOOKUP(A2,'" & RequisitionsSheetName & "'!A:B,2,0)", "=B2-I2", "=G2-I2", "=MIN(I2,G2-M2)", "=SUMIF('" & ReleasedOrdersSheetName & "'!A:A,A2,'" & ReleasedOrdersSheetName & "'!B:B)", "=G2-M2")
Sheets(MainSheetName).Range("C2").AutoFill Destination:=Range(Cells(2, 3), Cells(2, 6))

Sheets(RequisitionsSheetName).Range("A1:C1").Value = Array("Part Numbers", "Sum of Quantity", "Priority")

Sheets(ReleasedOrdersSheetName).Range("A1:B1").Value = Array("Part Numbers", "Lot Size", "Priority")

Sheets(MainSheetName).Range("A2").Select

End Sub
