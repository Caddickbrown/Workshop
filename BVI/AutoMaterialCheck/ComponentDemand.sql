SELECT
	mac.ORDER_NO AS 'SO Number',
	so.PART_NO AS 'Kit Number',
	mac.PART_NO AS 'Component Part Number',
	mac.LINE_ITEM_NO AS 'Line Number',
	mac.DATE_REQUIRED AS 'Date Needed',
	mac.QTY_REQUIRED AS 'Component Qty Required'

FROM IFS.SHOP_MATERIAL_ALLOC_TAB AS mac
LEFT JOIN IFS.SHOP_ORD_TAB AS so
	ON mac.ORDER_NO = so.ORDER_NO AND mac.CONTRACT = so.CONTRACT
WHERE
	so.CONTRACT = '2051'
	AND so.ROWSTATE IN ('Released')
;