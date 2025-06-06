-- Supplier for Purchased Part SQL Statement
SELECT 
	ipt.PART_NO As 'Part Number',
	sfpp.VENDOR_NO AS 'Supplier Code',
	CASE WHEN supp.NAME IS NULL
			THEN 'Manufactured'
		ELSE supp.NAME
		END AS 'Supplier Name',
	CASE WHEN ipt.LEAD_TIME_CODE = 'P'
			THEN 'Purchased'
		WHEN ipt.LEAD_TIME_CODE = 'M'
			THEN 'Manufactured'
		ELSE 'Error'
		END AS 'Part Type',
	CASE WHEN ipt.PLANNER_BUYER = 'MBANNER'
			THEN 'MB'
		WHEN ipt.PLANNER_BUYER = 'EGADOMSKA'
			THEN 'EG'
		WHEN ipt.PLANNER_BUYER = 'KDUFFILL'
			THEN 'KD'
		ELSE ipt.PLANNER_BUYER
		END AS 'Planner'
FROM IFS.INVENTORY_PART_TAB AS ipt
INNER JOIN  IFS.INVENTORY_PART_PLANNING_TAB AS ipp
	ON ipt.PART_NO = ipp.PART_NO AND ipt.CONTRACT = ipp.CONTRACT
LEFT JOIN  IFS.MANUF_STRUCTURE_HEAD_TAB AS ms
	ON ipt.PART_NO = ms.PART_NO AND ipt.CONTRACT = ms.CONTRACT
LEFT JOIN IFS.PURCHASE_PART_SUPPLIER_TAB as sfpp
	ON sfpp.CONTRACT = ipt.CONTRACT AND sfpp.PART_NO = ipt.PART_NO
LEFT JOIN IFS.SUPPLIER_INFO_TAB as supp
	ON sfpp.VENDOR_NO = supp.SUPPLIER_ID
WHERE ipt.CONTRACT = '2051'
AND (sfpp.PRIMARY_VENDOR IS NULL OR sfpp.PRIMARY_VENDOR = 'Y')
AND ms.EFF_PHASE_OUT_DATE IS NULL
ORDER BY ipt.PART_NO, ms.EFF_PHASE_IN_DATE DESC
;

-- # Changelog

-- ## [1.0.0] - 2024-08-15

-- ### Added

-- - Initial Commit

-- ### Changed

-- - Adjusted Part Type to look at LEAD_TIME_CODE
