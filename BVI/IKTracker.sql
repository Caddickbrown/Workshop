SELECT -- BVI Requisitions
	req.PROPOSAL_NO AS 'Order Number',
	'-' AS 'Priority Category',
	req.PART_NO AS 'Part No',
	ipt.PRIME_COMMODITY AS 'Comm Group',
	req.PROP_START_DATE AS '(Proposed) Start Date',
	req.PLAN_ORDER_REC AS 'Qty',
	'-' AS 'PO Number',
	'-' AS 'Operation',
	'BVI Req' AS 'Type'
FROM IFS.SHOP_ORDER_PROP_TAB AS req
INNER JOIN IFS.INVENTORY_PART_TAB AS ipt
    ON req.PART_NO = ipt.PART_NO AND req.CONTRACT = ipt.CONTRACT
WHERE req.CONTRACT = '2051'
AND req.ROWSTATE = 'ProposalCreated'
AND req.PROP_START_DATE <= GETDATE() + 150
AND ipt.PLANNER_BUYER IN ('2001','1001','TRAYS','SP_BIDFORD','4001','2051','1051','SUINST 4')

UNION

SELECT -- Malosa Requisitions
	req.PROPOSAL_NO AS 'Order Number',
	'-' AS 'Priority Category',
	req.PART_NO AS 'Part No',
	ipt.PRIME_COMMODITY AS 'Comm Group',
	req.PROP_START_DATE AS '(Proposed) Start Date',
	req.PLAN_ORDER_REC AS 'Qty',
	'-' AS 'PO Number',
	'-' AS 'Operation',
	'Malosa Req' AS 'Type'
FROM IFS.SHOP_ORDER_PROP_TAB AS req
INNER JOIN IFS.INVENTORY_PART_TAB AS ipt
    ON req.PART_NO = ipt.PART_NO AND req.CONTRACT = ipt.CONTRACT
WHERE req.CONTRACT = '2051'
AND req.ROWSTATE = 'ProposalCreated'
AND req.PROP_START_DATE <= GETDATE() + 150
AND LEFT(req.PART_NO,4) = 'MMSU'

UNION

SELECT -- Open Orders
	so.ORDER_NO AS 'Order Number',
	so.PRIORITY_CATEGORY AS 'Priority Category',
	so.PART_NO AS 'Part No',
	ipt.PRIME_COMMODITY AS 'Comm Group',
	so.REVISED_START_DATE AS '(Proposed) Start Date',
	(so.REVISED_QTY_DUE - so.QTY_COMPLETE) AS 'Qty', -- Qty Remaining
	pol.ORDER_NO AS 'PO Number',
	CASE WHEN so.ROWSTATE = 'Released' THEN 'Released' WHEN so.ROWSTATE = 'Planned' THEN 'Planned' WHEN pol.ORDER_NO IS NULL THEN 'Started' ELSE 'Steriliser' END AS 'Operation',
	'Released' AS 'Type'
FROM IFS.SHOP_ORD_TAB AS so
INNER JOIN IFS.INVENTORY_PART_TAB AS ipt
    ON so.PART_NO = ipt.PART_NO AND so.CONTRACT = ipt.CONTRACT
LEFT JOIN IFS.PURCHASE_ORDER_LINE_TAB AS pol
	ON so.ORDER_NO = pol.DEMAND_ORDER_NO
WHERE so.CONTRACT = '2051'
AND so.ROWSTATE IN ('Planned','Released','Started')
;