SELECT
req.PROPOSAL_NO AS Proposal, -- Proposal Number
req.PART_NO AS PartNo, -- Part Number
req.PROP_START_DATE AS ProposedStartDate, -- Proposed Start Date
req.REVISED_DUE_DATE AS RevisedDueDate, -- Expected Due Date
req.PLAN_ORDER_REC AS Qty -- Lot Size (Quantity)
FROM IFS.SHOP_ORDER_PROP_TAB AS req -- From Shop Order Requisitions Info
INNER JOIN IFS.INVENTORY_PART_TAB AS ipt -- Connect to Inventory Part Info
	ON req.PART_NO = ipt.PART_NO AND req.CONTRACT = ipt.CONTRACT -- Join on Part Number and Site
WHERE req.CONTRACT = '2051' -- Only on Site 2051 (Bidford)
AND req.ROWSTATE = 'ProposalCreated' -- Only Proposals
AND req.PROP_START_DATE < (GETDATE ()+63) -- Only the data for the next 9 weeks
AND ipt.PLANNER_BUYER IN ('3001','3801','5001'); -- Only Kit Data