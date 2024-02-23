
--WNK Unit Volume
SELECT a1."House", a1."ItemNumber", a1.ITCLS, a1."MasterUnitCubes", a1."OnHand", 0 as "Open PO"
FROM
(SELECT t1.STID AS "House", t1.ITNBR as "ItemNumber", t1.ITCLS, t1.B2Z95S as "MasterUnitCubes", t2.MOHTQ as "OnHand"
FROM AMFLIBW.ITMRVA t1  
LEFT JOIN AMFLIBW.ITEMBL t2 ON t1.ITNBR=t2.ITNBR and t1.STID=t2.HOUSE
WHERE t1.STID IN ('35') AND t1.ITCLS like 'Z%' and t1.ITCLS not like '%K' AND t1.ITNBR NOT LIKE 'A%'
) a1



-- Ashton Unit Volume

Select a.HOUSE,a.Item_Number,a.ITCLS, a.Master_Unit_Cube, a.On_Hand, b.Open_PO,
(CASE 
	WHEN a.Item_Number LIKE 'A%' THEN 'CG'
	WHEN a.Item_Number LIKE 'B%' THEN 'CG'
	WHEN a.Item_Number LIKE 'D%' THEN 'CG'
	WHEN a.Item_Number LIKE 'E%' THEN 'CG'
	WHEN a.Item_Number LIKE 'H%' THEN 'CG'
	WHEN a.Item_Number LIKE 'L%' THEN 'CG'
	WHEN a.Item_Number LIKE 'M%' THEN 'CG'
	WHEN a.Item_Number LIKE 'P%' THEN 'CG'
	WHEN a.Item_Number LIKE 'Q%' THEN 'CG'
	WHEN a.Item_Number LIKE 'R%' THEN 'CG'
	WHEN a.Item_Number LIKE 'T%' THEN 'CG'
	WHEN a.Item_Number LIKE 'W%' THEN 'CG'
	WHEN a.Item_Number LIKE 'X%' THEN 'CG'
	WHEN a.Item_Number LIKE 'Y%' THEN 'CG'
	WHEN a.Item_Number LIKE 'Z%' THEN 'CG'
	ELSE 'UPH' END) AS Product
from 
(SELECT t2.HOUSE,trim(t4.ITNBR) as Item_Number, t2.ITNBR,t2.ITCLS, t4.B2Z95S as Master_Unit_Cube, t2.MOHTQ as On_Hand  FROM AFILELIB.ITBEXT t1, AMFLIBA.ITEMBL t2, AFILELIB.ITMEXT t3, AMFLIBA.ITMRVA t4
WHERE t1.ITNBR = t4.ITNBR AND t2.HOUSE = t1.HOUSE AND t2.ITNBR = t1.ITNBR AND t2.ITNBR = t4.ITNBR AND t2.ITCLS = t4.ITCLS AND t3.ITNBR = t1.ITNBR AND t3.ITNBR = t2.ITNBR AND t3.ITNBR = t4.ITNBR AND t4.STID = t1.HOUSE AND ((t1.HOUSE='335')) and t2.ITCLS like 'Z%' AND t2.ITCLS NOT LIKE '%K') a 
left join 
(SELECT trim(t1.ITNBR) as Item_Number, sum(t1.QTYOR) as Open_PO FROM AMFLIBA.POITEM t1, AMFLIBA.POMAST t2 WHERE t1.ORDNO = t2.ORDNO AND t2.HOUSE = t1.HOUSE AND (t1.HOUSE='335') AND PSTTS IN ('10','20','30') group by t1.ITNBR ORDER BY t1.itnbr) b
on a.Item_Number = b.Item_Number
WHERE (CASE 
	WHEN a.Item_Number LIKE 'A%' THEN 'CG'
	WHEN a.Item_Number LIKE 'B%' THEN 'CG'
	WHEN a.Item_Number LIKE 'D%' THEN 'CG'
	WHEN a.Item_Number LIKE 'E%' THEN 'CG'
	WHEN a.Item_Number LIKE 'H%' THEN 'CG'
	WHEN a.Item_Number LIKE 'L%' THEN 'CG'
	WHEN a.Item_Number LIKE 'M%' THEN 'CG'
	WHEN a.Item_Number LIKE 'P%' THEN 'CG'
	WHEN a.Item_Number LIKE 'Q%' THEN 'CG'
	WHEN a.Item_Number LIKE 'R%' THEN 'CG'
	WHEN a.Item_Number LIKE 'T%' THEN 'CG'
	WHEN a.Item_Number LIKE 'W%' THEN 'CG'
	WHEN a.Item_Number LIKE 'X%' THEN 'CG'
	WHEN a.Item_Number LIKE 'Y%' THEN 'CG'
	WHEN a.Item_Number LIKE 'Z%' THEN 'CG'
	ELSE 'UPH' END) LIKE 'UPH'


