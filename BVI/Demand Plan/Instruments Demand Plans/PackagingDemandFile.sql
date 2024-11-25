DECLARE @WeeksOut INT;
DECLARE @SterileWeeksOut INT;
 
SET @WeeksOut = 4;
SET @SterileWeeksOut = 4;

SELECT
	--*,
	req.PROPOSAL_NO AS 'Proposal/Order No',
	CASE WHEN TRY_CAST(req.PART_NO AS INT) IS NOT NULL THEN CAST(req.PART_NO AS INT) ELSE req.PART_NO END AS 'Part No',
	ipt.PLANNER_BUYER AS 'Planner',
	ipt.PRIME_COMMODITY AS 'Comm Group',
	req.PROP_START_DATE AS '"Start" Date',
	req.PLAN_ORDER_REC AS 'Qty',
	CASE WHEN LEFT(req.PART_NO,4) = 'PLAN' THEN 'Plan Order' ELSE 'Normal Order' END AS 'Forecast Order'
	
FROM IFS.SHOP_ORDER_PROP_TAB AS req
INNER JOIN IFS.INVENTORY_PART_TAB AS ipt ON req.PART_NO = ipt.PART_NO AND req.CONTRACT = ipt.CONTRACT
--INNER JOIN IFS.ROUTING_OPERATION_TAB AS hrs ON req.PART_NO = hrs.PART_NO AND ipt.CONTRACT = hrs.CONTRACT

WHERE req.CONTRACT = '2051'
AND req.ROWSTATE = 'ProposalCreated'
AND ipt.PLANNER_BUYER IN ('2001','1001','TRAYS','SP_BIDFORD','4001','2051','1051','SUINST 4')
--AND req.PART_NO IN ('590389')
AND req.PROP_START_DATE <= DATEADD(DAY, (@WeeksOut+1)*7 - DATEPART(WEEKDAY, GETDATE()), GETDATE()) -- This will round to before the Monday after however many weeks out
--AND req.PROP_START_DATE <= '09/01/2024' -- Before specific Date (mm/dd/yyyy)
AND ipt.PRIME_COMMODITY NOT IN ('XSTAR','SAFET','MVRB')
AND req.PART_NO LIKE '8%'

UNION

SELECT
	--*,
	req.PROPOSAL_NO AS 'Proposal/Order No',
	CASE WHEN TRY_CAST(req.PART_NO AS INT) IS NOT NULL THEN CAST(req.PART_NO AS INT) ELSE req.PART_NO END AS 'Part No',
	ipt.PLANNER_BUYER AS 'Planner',
	ipt.PRIME_COMMODITY AS 'Comm Group',
	req.PROP_START_DATE AS '"Start" Date',
	req.PLAN_ORDER_REC AS 'Qty',
	CASE WHEN LEFT(req.PART_NO,4) = 'PLAN' THEN 'Plan Order' ELSE 'Normal Order' END AS 'Forecast Order'
	
FROM IFS.SHOP_ORDER_PROP_TAB AS req
INNER JOIN IFS.INVENTORY_PART_TAB AS ipt ON req.PART_NO = ipt.PART_NO AND req.CONTRACT = ipt.CONTRACT
--INNER JOIN IFS.ROUTING_OPERATION_TAB AS hrs ON req.PART_NO = hrs.PART_NO AND ipt.CONTRACT = hrs.CONTRACT

WHERE req.CONTRACT = '2051'
AND req.ROWSTATE = 'ProposalCreated'
AND ipt.PLANNER_BUYER IN ('2001','1001','TRAYS','SP_BIDFORD','4001','2051','1051','SUINST 4')
--AND req.PART_NO IN ('590389')
AND req.PROP_START_DATE <= DATEADD(DAY, (@SterileWeeksOut+1)*7 - DATEPART(WEEKDAY, GETDATE()), GETDATE()) -- This will round to before the Monday after however many weeks out
--AND req.PROP_START_DATE <= '09/01/2024' -- Before specific Date (mm/dd/yyyy)
AND ipt.PRIME_COMMODITY NOT IN ('XSTAR','SAFET','MVRB')
AND req.PART_NO NOT LIKE '8%'

ORDER BY req.PROP_START_DATE
;

-- # Changelog

-- ## [2.0.0] - 2024-11-25

-- ### Added

-- - Variable for Weeks Out
-- - Variable for Sterile Weeks Out (Including new Query to search separately)
-- - Initial Commit

-- ### Changed

-- - Explaination for Days out filter