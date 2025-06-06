SELECT
    PROPOSAL_NO AS ProposalNo,
    PART_NO AS PartNo,
    PROP_START_DATE AS PlanStartDate,
    PLAN_ORDER_REC AS PlanQty
FROM IFS.SHOP_ORDER_PROP_TAB AS req
WHERE CONTRACT = '2051'
AND ROWSTATE = 'ProposalCreated'
--  AND PART_NO NOT LIKE 'NS%'
AND PROP_START_DATE <= (GETDATE ()+63);