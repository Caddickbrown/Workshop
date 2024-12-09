DECLARE @WeeksOut INT;
DECLARE @SterileWeeksOut INT;
 
SET @WeeksOut = 4;
SET @SterileWeeksOut = 8;
 
SELECT
    req.PROPOSAL_NO AS 'Proposal/Order No',
    req.PART_NO AS 'Part No',
    ipt.PLANNER_BUYER AS 'Planner',
    ipt.PRIME_COMMODITY AS 'Comm Group',
    req.PROP_START_DATE AS '"Start" Date',
    req.PLAN_ORDER_REC AS 'Qty',
    CASE 
        WHEN RIGHT(req.PART_NO, 1) = 'S' THEN 'Sterile' 
        ELSE 'Non-Sterile' 
    END AS 'Sterility',
    'Requisition' AS 'Order Type'
FROM 
    IFS.SHOP_ORDER_PROP_TAB AS req
INNER JOIN 
    IFS.INVENTORY_PART_TAB AS ipt
    ON req.PART_NO = ipt.PART_NO AND req.CONTRACT = ipt.CONTRACT
WHERE 
    req.CONTRACT = '2051'
    AND req.ROWSTATE = 'ProposalCreated'
    AND LEFT(req.PART_NO, 4) = 'MMSU'
    AND req.PROP_START_DATE <= DATEADD(DAY, (@WeeksOut+1)*7 - DATEPART(WEEKDAY, GETDATE()), GETDATE()) -- This will round to before the Monday after however many weeks out
    AND req.PART_NO NOT LIKE '%S'

UNION

SELECT
    req.PROPOSAL_NO AS 'Proposal/Order No',
    req.PART_NO AS 'Part No',
    ipt.PLANNER_BUYER AS 'Planner',
    ipt.PRIME_COMMODITY AS 'Comm Group',
    req.PROP_START_DATE AS '"Start" Date',
    req.PLAN_ORDER_REC AS 'Qty',
    CASE 
        WHEN RIGHT(req.PART_NO, 1) = 'S' THEN 'Sterile' 
        ELSE 'Non-Sterile' 
    END AS 'Sterility',
    'Requisition' AS 'Order Type'
FROM 
    IFS.SHOP_ORDER_PROP_TAB AS req
INNER JOIN 
    IFS.INVENTORY_PART_TAB AS ipt 
    ON req.PART_NO = ipt.PART_NO AND req.CONTRACT = ipt.CONTRACT
WHERE 
    req.CONTRACT = '2051'
    AND req.ROWSTATE = 'ProposalCreated'
    AND LEFT(req.PART_NO, 4) = 'MMSU'
    AND req.PROP_START_DATE <= DATEADD(DAY, (@SterileWeeksOut+1)*7 - DATEPART(WEEKDAY, GETDATE()), GETDATE()) -- Before the Monday two weeks after next
    AND req.PART_NO LIKE '%S'
ORDER BY 
    req.PROP_START_DATE
;

-- # Changelog

-- ## [2.0.0] - 2024-11-25

-- ### Added

-- - Variable for Weeks Out
-- - Variable for Sterile Weeks Out (Including new Query to search separately)
-- - Initial Commit

-- ### Changed

-- - Explaination for Days out filter