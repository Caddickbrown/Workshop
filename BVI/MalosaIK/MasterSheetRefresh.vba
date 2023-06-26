' ## To Do
' - [ ] Auto Pull Exported Data sheets
' - [ ] Formatting Reset 
' - [ ] 

Sub MasterSheetRefresh()

Dim MainSheetName As String, RequisitionsSheetName As String, ReleasedOrdersSheetName As String, IPISSheetName As String, PartNumberCalc As String, TotalRawMaterialQtyCalc As String, AMCOCalc As String, B1StockCalc As String, RMMaterialCalc As String, TotalReqForWeekCalc As String, RMShortageCalc As String, B1ShortageCalc As String, QuickReleaseCalc As String, ReleasedSOCalc As String, NetUsableRMCalc As String

'Tab Names
MainSheetName = "Locations"
RequisitionsSheetName = "Requisition Demand"
ReleasedOrdersSheetName = "Released Shop Orders"
IPISSheetName = "IPIS"

'Formulas
PartNumberCalc = "='" & RequisitionsSheetName & "'!A2"
TotalRawMaterialQtyCalc = "=SUMIF(INDEX(" & IPISSheetName & "!$A:$BZ,0,MATCH(""Part No""," & IPISSheetName & "!$A$1:$BZ$1,0)),LEFT(A2,8)&""A"",INDEX(" & IPISSheetName & "!$A:$BZ,0,MATCH(""On Hand Qty""," & IPISSheetName & "!$A$1:$BZ$1,0)))"
AMCOCalc = "=SUMIFS(INDEX(" & IPISSheetName & "!$A:$BZ,0,MATCH(""On Hand Qty""," & IPISSheetName & "!$A$1:$BZ$1,0)),INDEX(" & IPISSheetName & "!$A:$BZ,0,MATCH(""Warehouse""," & IPISSheetName & "!$A$1:$BZ$1,0)),C$1,INDEX(" & IPISSheetName & "!$A:$BZ,0,MATCH(""Part No""," & IPISSheetName & "!$A$1:$BZ$1,0)),LEFT($A2,8)&""A"")"
B1StockCalc = "=E2+F2"
RMMaterialCalc = "=CONCATENATE(LEFT(A2,8),""A"")"
TotalReqForWeekCalc = "=VLOOKUP(A2,'" & RequisitionsSheetName & "'!A:B,2,0)"
RMShortageCalc = "=B2-I2"
B1ShortageCalc = "=G2-I2"
QuickReleaseCalc = "=MIN(I2,G2-M2)"
ReleasedSOCalc = "=SUMIF('" & ReleasedOrdersSheetName & "'!A:A,A2,'" & ReleasedOrdersSheetName & "'!B:B)"
NetUsableRMCalc = "=G2-M2"

Sheets(MainSheetName).Rows("3:1048576").Delete Shift:=xlUp
Sheets(RequisitionsSheetName).Rows("2:1048576").ClearContents
Sheets(ReleasedOrdersSheetName).Rows("2:1048576").ClearContents
Sheets(IPISSheetName).Cells.ClearContents

Sheets(MainSheetName).Range("A1:N1").Value = Array("Part Number", "Total Raw Material Qty", "AMCO", "GOODS-IN", "INST&KNIVES", "CENTRAL-STORES", "B1 Stock", "RM Material", "Total Req For Week", "RM Shortage", "B1 Shortage", "Quick Release", "Released SO", "Net Usable RM")
Sheets(MainSheetName).Range("A2:N2").Formula2 = Array(PartNumberCalc,TotalRawMaterialQtyCalc, AMCOCalc, "", "", "", B1StockCalc, RMMaterialCalc, TotalReqForWeekCalc, RMShortageCalc, B1ShortageCalc, QuickReleaseCalc, ReleasedSOCalc, NetUsableRMCalc)
Sheets(MainSheetName).Range("C2").AutoFill Destination:=Range(Cells(2, 3), Cells(2, 6))

Sheets(RequisitionsSheetName).Range("A1:C1").Value = Array("Part Numbers", "Sum of Quantity", "Priority")

Sheets(ReleasedOrdersSheetName).Range("A1:B1").Value = Array("Part Numbers", "Lot Size", "Priority")

Sheets(MainSheetName).Range("A2").Select

End Sub
