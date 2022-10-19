SELECT
so.ORDER_NO, -- Shop Order Number
so.PART_NO, -- Part Number
so.REVISED_START_DATE, -- Current Start Date
so.REVISED_QTY_DUE -- Currrent Qty
FROM IFS.SHOP_ORD_TAB AS so -- From the Shop Order Information
INNER JOIN IFS.INVENTORY_PART_TAB AS ipt -- Connect to Inventory Part Info
	ON so.PART_NO = ipt.PART_NO AND so.CONTRACT = ipt.CONTRACT -- Join on Part Number and Site
WHERE so.CONTRACT = '2051' -- Only on Site 2051 (Bidford)
AND so.ROWSTATE = 'Released' -- Only Released Orders
AND so.REVISED_START_DATE < (GETDATE ()+63) -- Only the data for the next 9 weeks
AND ipt.PLANNER_BUYER IN ('3001','3801','5001'); -- Only Kit Data