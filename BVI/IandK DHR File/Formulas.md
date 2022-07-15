
## Part No (Array)
=TRIM(SORT(UNIQUE(FILTER('SO Reqs'!C2:C1048576,'SO Reqs'!C2:C1048576<>""),0,0),,1))

## Start Date (Array)
=INDEX('SO Reqs'!O:O,MATCH(SORT(UNIQUE(FILTER('SO Reqs'!C2:C1048576,'SO Reqs'!C2:C1048576<>""),0,0),,1),'SO Reqs'!C:C,0))

## Plan (Array)
=INDEX('SO Reqs'!W:W,MATCH(SORT(UNIQUE(FILTER('SO Reqs'!C2:C1048576,'SO Reqs'!C2:C1048576<>""),0,0),,1),'SO Reqs'!C:C,0))

## Manu Part (Array)
=IFERROR(INDEX(Data!B2:B1570,MATCH(TRIM(SORT(UNIQUE(FILTER('SO Reqs'!C2:C1048576,'SO Reqs'!C2:C1048576<>""),0,0),,1)),Data!A2:A1570,0)),"-")

## Stock (Array)
=SUMIF(IPIS!B:B,INDEX(Data!B2:B1570,MATCH(TRIM(SORT(UNIQUE(FILTER('SO Reqs'!C2:C1048576,'SO Reqs'!C2:C1048576<>""),0,0),,1)),Data!A2:A1570,0)),IPIS!D:D)

## Sterile
=SUM(IFERROR(INDEX(ManStru!V:V,MATCH(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)&A2,ManStru!C:C&ManStru!N:N,0))*IF(LEFT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),2)="58",IF(LEFT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),4)="5857",0,IF(LEFT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),3)="588",0,IF(LEFT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),3)="589",0,IF(LEFT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),3)="590",0,SUMIFS('Parent Reqs'!K:K,'Parent Reqs'!C:C,UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)))))),0),0))

## Sterile US
=SUM(IFERROR(INDEX(ManStru!V:V,MATCH(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)&A2,ManStru!C:C&ManStru!N:N,0))*IF(RIGHT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),2)="US",SUMIFS('Parent Reqs'!K:K,'Parent Reqs'!C:C,UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)),0),0))

## Non-Ster
=SUM(IFERROR(INDEX(ManStru!V:V,MATCH(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)&A2,ManStru!C:C&ManStru!N:N,0))*IF(LEFT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),4)="5857",SUMIFS('Parent Reqs'!K:K,'Parent Reqs'!C:C,UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)),0),0))

## Kits
=SUM(IFERROR(INDEX(ManStru!V:V,MATCH(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)&A2,ManStru!C:C&ManStru!N:N,0))*IF(LEFT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),2)="80",SUMIFS('Parent Reqs'!K:K,'Parent Reqs'!C:C,UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)),IF(LEFT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),3)="588",SUMIFS('Parent Reqs'!K:K,'Parent Reqs'!C:C,UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)),IF(LEFT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),3)="589",SUMIFS('Parent Reqs'!K:K,'Parent Reqs'!C:C,UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)),IF(LEFT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),3)="590",SUMIFS('Parent Reqs'!K:K,'Parent Reqs'!C:C,UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)),IF(LEFT(UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0),2)="NS",SUMIFS('Parent Reqs'!K:K,'Parent Reqs'!C:C,UNIQUE(FILTER(Data!$A$2:$A$1570,Data!$B$2:$B$1570=A2),0,0)),0))))),0))
