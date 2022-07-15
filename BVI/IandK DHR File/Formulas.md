
Need to convert the below to PowerQuery of something of the sort

Array formulas are not suitable for largish amounts of data, but there are suitable alternatives: changing to VBA won't speed up the process if you still use array formulas - they are structurally always going to be very slow for large amounts of data. Instead of array formulas, use efficient database type functionality. Maybe pivot tables or query tables or (via VBA) recordsets. These don't use worksheet formulas. There are examples in old threads where array formulas took more then an hour but database type approaches a fraction of a second.

## Variable Column Lookup

## Part No (Array) Variable Lookup
=TRIM(SORT(UNIQUE(FILTER(INDEX('SO Reqs'!A:DZ,0,MATCH("Part No",'SO Reqs'!A$1:DZ$1,0)),OR(INDEX('SO Reqs'!A:DZ,0,MATCH("Part No",'SO Reqs'!A$1:DZ$1,0))<>"",INDEX('SO Reqs'!A:DZ,0,MATCH("Part No",'SO Reqs'!A$1:DZ$1,0))<>"Part No")),0,0),,1))

## Start Date (Array) Variable Lookup
=INDEX(INDEX('SO Reqs'!A:DZ,0,MATCH("Proposed Start Date",'SO Reqs'!A$1:DZ$1,0)),MATCH(SORT(UNIQUE(FILTER(INDEX('SO Reqs'!A:DZ,0,MATCH("Part No",'SO Reqs'!A$1:DZ$1,0)),OR(INDEX('SO Reqs'!A:DZ,0,MATCH("Part No",'SO Reqs'!A$1:DZ$1,0))<>"",INDEX('SO Reqs'!A:DZ,0,MATCH("Part No",'SO Reqs'!A$1:DZ$1,0))<>"Part No")),0,0),,1),INDEX('SO Reqs'!A:DZ,0,MATCH("Part No",'SO Reqs'!A$1:DZ$1,0)),0))

## Plan (Array) Variable Lookup
=INDEX(INDEX('SO Reqs'!A:DZ,0,MATCH("Planner",'SO Reqs'!A$1:DZ$1,0)),MATCH(SORT(UNIQUE(FILTER(INDEX('SO Reqs'!A:DZ,0,MATCH("Part No",'SO Reqs'!A$1:DZ$1,0)),OR(INDEX('SO Reqs'!A:DZ,0,MATCH("Part No",'SO Reqs'!A$1:DZ$1,0))<>"",INDEX('SO Reqs'!A:DZ,0,MATCH("Part No",'SO Reqs'!A$1:DZ$1,0))<>"Part No")),0,0),,1),INDEX('SO Reqs'!A:DZ,0,MATCH("Part No",'SO Reqs'!A$1:DZ$1,0)),0))

## Manu Part (Array) Variable Lookup
=IFERROR(INDEX(Data!B$2:B$1570,MATCH(TRIM(SORT(UNIQUE(FILTER(INDEX('SO Reqs'!A:DZ,0,MATCH("Part No",'SO Reqs'!A$1:DZ$1,0)),OR(INDEX('SO Reqs'!A:DZ,0,MATCH("Part No",'SO Reqs'!A$1:DZ$1,0))<>"",INDEX('SO Reqs'!A:DZ,0,MATCH("Part No",'SO Reqs'!A$1:DZ$1,0))<>"Part No")),0,0),,1)),Data!A$2:A$1570,0)),"-")

## Stock Variable Lookup
=SUMIF(INDEX(IPIS!A:DZ,0,MATCH("Part No",IPIS!A$1:DZ$1,0)),INDEX(Data!B$2:B$1570,MATCH(A2,Data!A$2:A$1570,0)),INDEX(IPIS!A:DZ,0,MATCH("On Hand Qty",IPIS!A$1:DZ$1,0)))-SUMIF(INDEX('Released Shop Orders'!A:DZ,0,MATCH("Part No",'Released Shop Orders'!A$1:DZ$1,0)),A2,INDEX('Released Shop Orders'!A:DZ,0,MATCH("Lot Size",'Released Shop Orders'!A$1:DZ$1,0)))

## Sterile Variable Lookup
=SUM(IFERROR(INDEX(ManStru!V:V,MATCH(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)&A2,ManStru!C:C&ManStru!N:N,0))*IF(LEFT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),2)="58",IF(LEFT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),4)="5857",0,IF(LEFT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),3)="588",0,IF(LEFT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),3)="589",0,IF(LEFT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),3)="590",0,SUMIFS(INDEX('Parent Reqs'!A:DZ,0,MATCH("Quantity",'Parent Reqs'!A$1:DZ$1,0)),INDEX('Parent Reqs'!A:DZ,0,MATCH("Part No",'Parent Reqs'!A$1:DZ$1,0)),UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)))))),0),0))

## Sterile US Variable Lookup
=SUM(IFERROR(INDEX(ManStru!V:V,MATCH(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)&A2,ManStru!C:C&ManStru!N:N,0))*IF(RIGHT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),2)="US",SUMIFS(INDEX('Parent Reqs'!A:DZ,0,MATCH("Quantity",'Parent Reqs'!A$1:DZ$1,0)),INDEX('Parent Reqs'!A:DZ,0,MATCH("Part No",'Parent Reqs'!A$1:DZ$1,0)),UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)),0),0))

## Non-Ster Variable Lookup
=SUM(IFERROR(INDEX(ManStru!V:V,MATCH(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)&A2,ManStru!C:C&ManStru!N:N,0))*IF(LEFT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),4)="5857",SUMIFS(INDEX('Parent Reqs'!A:DZ,0,MATCH("Quantity",'Parent Reqs'!A$1:DZ$1,0)),INDEX('Parent Reqs'!A:DZ,0,MATCH("Part No",'Parent Reqs'!A$1:DZ$1,0)),UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)),0),0))

## Kits Variable Lookup
=SUM(IFERROR(INDEX(ManStru!V:V,MATCH(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)&A2,ManStru!C:C&ManStru!N:N,0))*IF(LEFT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),2)="80",SUMIFS(INDEX('Parent Reqs'!A:DZ,0,MATCH("Quantity",'Parent Reqs'!A$1:DZ$1,0)),INDEX('Parent Reqs'!A:DZ,0,MATCH("Part No",'Parent Reqs'!A$1:DZ$1,0)),UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)),IF(LEFT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),3)="588",SUMIFS(INDEX('Parent Reqs'!A:DZ,0,MATCH("Quantity",'Parent Reqs'!A$1:DZ$1,0)),INDEX('Parent Reqs'!A:DZ,0,MATCH("Part No",'Parent Reqs'!A$1:DZ$1,0)),UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)),IF(LEFT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),3)="589",SUMIFS(INDEX('Parent Reqs'!A:DZ,0,MATCH("Quantity",'Parent Reqs'!A$1:DZ$1,0)),INDEX('Parent Reqs'!A:DZ,0,MATCH("Part No",'Parent Reqs'!A$1:DZ$1,0)),UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)),IF(LEFT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),3)="590",SUMIFS(INDEX('Parent Reqs'!A:DZ,0,MATCH("Quantity",'Parent Reqs'!A$1:DZ$1,0)),INDEX('Parent Reqs'!A:DZ,0,MATCH("Part No",'Parent Reqs'!A$1:DZ$1,0)),UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)),IF(LEFT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),2)="NS",SUMIFS(INDEX('Parent Reqs'!A:DZ,0,MATCH("Quantity",'Parent Reqs'!A$1:DZ$1,0)),INDEX('Parent Reqs'!A:DZ,0,MATCH("Part No",'Parent Reqs'!A$1:DZ$1,0)),UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)),0))))),0))

# Ex.

## Part No (Array) (Fixed Columns)
=TRIM(SORT(UNIQUE(FILTER('SO Reqs'!C2:C1048576,'SO Reqs'!C2:C1048576<>""),0,0),,1))

## Start Date (Array) (Fixed Columns)
=INDEX('SO Reqs'!O:O,MATCH(SORT(UNIQUE(FILTER('SO Reqs'!C2:C1048576,'SO Reqs'!C2:C1048576<>""),0,0),,1),'SO Reqs'!C:C,0))

## Plan (Array) (Fixed Columns)
=INDEX('SO Reqs'!W:W,MATCH(SORT(UNIQUE(FILTER('SO Reqs'!C2:C1048576,'SO Reqs'!C2:C1048576<>""),0,0),,1),'SO Reqs'!C:C,0))

## Manu Part (Array) (Fixed Columns)
=IFERROR(INDEX(Data!B2:B1570,MATCH(TRIM(SORT(UNIQUE(FILTER('SO Reqs'!C2:C1048576,'SO Reqs'!C2:C1048576<>""),0,0),,1)),Data!A2:A1570,0)),"-")

## Stock
=SUMIF(IPIS!B:B,INDEX(Data!B2:B1570,MATCH(A2,Data!A2:A1570,0)),IPIS!D:D)-SUMIF('Released Shop Orders'!D:D,A2,'Released Shop Orders'!M:M)

## Sterile (Fixed Columns)
=SUM(IFERROR(INDEX(ManStru!V:V,MATCH(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)&A2,ManStru!C:C&ManStru!N:N,0))*IF(LEFT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),2)="58",IF(LEFT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),4)="5857",0,IF(LEFT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),3)="588",0,IF(LEFT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),3)="589",0,IF(LEFT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),3)="590",0,SUMIFS('Parent Reqs'!K:K,'Parent Reqs'!C:C,UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)))))),0),0))

## Sterile US (Fixed Columns)
=SUM(IFERROR(INDEX(ManStru!V:V,MATCH(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)&A2,ManStru!C:C&ManStru!N:N,0))*IF(RIGHT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),2)="US",SUMIFS('Parent Reqs'!K:K,'Parent Reqs'!C:C,UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)),0),0))

## Non-Ster (Fixed Columns)
=SUM(IFERROR(INDEX(ManStru!V:V,MATCH(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)&A2,ManStru!C:C&ManStru!N:N,0))*IF(LEFT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),4)="5857",SUMIFS('Parent Reqs'!K:K,'Parent Reqs'!C:C,UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)),0),0))

## Kits (Fixed Columns)
=SUM(IFERROR(INDEX(ManStru!V:V,MATCH(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)&A2,ManStru!C:C&ManStru!N:N,0))*IF(LEFT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),2)="80",SUMIFS('Parent Reqs'!K:K,'Parent Reqs'!C:C,UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)),IF(LEFT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),3)="588",SUMIFS('Parent Reqs'!K:K,'Parent Reqs'!C:C,UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)),IF(LEFT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),3)="589",SUMIFS('Parent Reqs'!K:K,'Parent Reqs'!C:C,UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)),IF(LEFT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),3)="590",SUMIFS('Parent Reqs'!K:K,'Parent Reqs'!C:C,UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)),IF(LEFT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),2)="NS",SUMIFS('Parent Reqs'!K:K,'Parent Reqs'!C:C,UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)),0))))),0))

## Stock (Array) (Doesn't take away released stock)
=SUMIF(IPIS!B:B,INDEX(Data!B2:B1570,MATCH(TRIM(SORT(UNIQUE(FILTER('SO Reqs'!C2:C1048576,'SO Reqs'!C2:C1048576<>""),0,0),,1)),Data!A2:A1570,0)),IPIS!D:D)