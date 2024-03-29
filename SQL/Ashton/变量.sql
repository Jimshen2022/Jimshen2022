DECLARE BS INT   --- BS 为期初库存 
SELECT T1."BeginningQty" INTO BS
FROM (
SELECT 
SUM(CASE WHEN "Type" IN ('CurrentStock') THEN QTY ELSE 0 END) AS "CurrentStockQty",
SUM(CASE WHEN "Type" IN ('IN') THEN QTY ELSE 0 END) AS "IN_QTY",
SUM(CASE WHEN "Type" IN ('OUT') THEN QTY ELSE 0 END) AS "OUT_QTY",
SUM(CASE WHEN "Type" IN ('CurrentStock') THEN "AMT($USD)" ELSE 0 END) AS "CurrentStockAMT",
SUM(CASE WHEN "Type" IN ('IN') THEN "AMT($USD)" ELSE 0 END) AS "IN_AMT",
SUM(CASE WHEN "Type" IN ('OUT') THEN "AMT($USD)" ELSE 0 END) AS "OUT_AMT",
SUM(CASE WHEN "Type" IN ('CurrentStock') THEN QTY ELSE 0 END) + 
SUM(CASE WHEN "Type" IN ('OUT') THEN QTY ELSE 0 END) - SUM(CASE WHEN "Type" IN ('IN') THEN QTY ELSE 0 END) as "BeginningQty",
SUM(CASE WHEN "Type" IN ('CurrentStock') THEN "AMT($USD)" ELSE 0 END) + 
SUM(CASE WHEN "Type" IN ('OUT') THEN "AMT($USD)" ELSE 0 END) - SUM(CASE WHEN "Type" IN ('IN') THEN "AMT($USD)" ELSE 0 END) as "BeginningAMT"
FROM (
SELECT x.WEEK,x."Type", SUM(x.MOHTQ) QTY, SUM(x."AMT($USD)") "AMT($USD)"
FROM (
-- CurrentOnHand
SELECT a.ITNBR,a.ITCLS, a.MOHTQ, b.PAMNT "UP($USD)",
(CASE 
    WHEN b.PAMNT IS NULL AND a.ITCLS NOT LIKE 'Z%' THEN 0
	WHEN b.PAMNT IS NULL AND a.ITCLS LIKE 'Z%' THEN a.MOHTQ*150
	ELSE a.MOHTQ*b.PAMNT END) as "AMT($USD)", 
'CurrentStock' AS "Type",
WEEK(CURRENT DATE) AS WEEK 
FROM 
(SELECT T1.ITNBR,T1.HOUSE,T1.ITCLS,T1.MOHTQ   
FROM AMFLIBA.ITEMBL T1, AMFLIBA.ITMRVA T2, AMFLIBA.WHSMST T3 
WHERE  T2.ITCLS = T1.ITCLS AND T2.ITNBR = T1.ITNBR AND T2.STID = T3.STID AND T1.HOUSE = T3.WHID AND ((T1.HOUSE='335') AND (T1.MOHTQ<>0))
ORDER BY T1.ITNBR) as a
LEFT JOIN (SELECT PRICE.PITEM, PRICE.PAMNT FROM AFILELIB.PRICE PRICE WHERE PRICE.PRICCD='FOBARC' ORDER BY PRICE.PITEM) as b 
ON a.ITNBR = b.PITEM 
) AS x
GROUP BY x.WEEK,x."Type"

UNION ALL

SELECT z.WEEK,z."Type",SUM(z.TRQTY) QTY, SUM(z."AMT($USD)") "AMT($USD)"
FROM(
-- get transactoins in and out  
SELECT X.ITNBR, X.ITCLS, X.TRQTY, Y.PAMNT "UP($USD)",X.WEEK,
(CASE 
    WHEN Y.PAMNT IS NULL AND X.ITCLS NOT LIKE 'Z%' THEN 0
	WHEN Y.PAMNT IS NULL AND X.ITCLS LIKE 'Z%' THEN X.TRQTY*150
	ELSE X.TRQTY*Y.PAMNT END) as "AMT($USD)",
(CASE 
		WHEN X.TCODE IN ('RP') THEN 'IN'
		WHEN X.TCODE IN ('SA') THEN 'OUT'
		ELSE 'CHECK' END ) as "Type"
		
FROM
(
SELECT T1.TCODE, T1.ITNBR, T1.HOUSE, T2.ITCLS, T1.TRQTY,
WEEK(DATE('20'||SUBSTR(CHAR(T1.UPDDT),2,2)||'-'||SUBSTR(CHAR(T1.UPDDT),4,2)||'-'||SUBSTR(CHAR(T1.UPDDT),6,2))) AS WEEK
FROM AMFLIBA.IMHIST T1, AMFLIBA.ITMRVA T2, AMFLIBA.WHSMST T3
WHERE T2.ITNBR = T1.ITNBR AND T2.STID = T3.STID AND T1.HOUSE = T3.WHID AND T1.HOUSE='335' 
AND T1.UPDDT BETWEEN '1210101' AND CHAR('1'||VARCHAR_FORMAT(current date,'YYMMDD'))
AND T1.TRQTY<>0 AND T1.TCODE in ('RP','SA') 
ORDER BY T1.ITNBR 
) X
LEFT JOIN (SELECT PRICE.PITEM, PRICE.PAMNT FROM AFILELIB.PRICE PRICE WHERE PRICE.PRICCD='FOBARC' ORDER BY PRICE.PITEM) as Y 
ON X.ITNBR = Y.PITEM
) AS z
GROUP BY z.WEEK,z."Type") A
) 


SELECT a.ITNBR,a.ITCLS, a.MOHTQ + BS, b.PAMNT "UP($USD)",
(CASE 
    WHEN b.PAMNT IS NULL AND a.ITCLS NOT LIKE 'Z%' THEN 0
	WHEN b.PAMNT IS NULL AND a.ITCLS LIKE 'Z%' THEN a.MOHTQ*150
	ELSE a.MOHTQ*b.PAMNT END) as "AMT($USD)", 
'CurrentStock' AS "Type",
WEEK(CURRENT DATE) AS WEEK 
FROM 
(SELECT T1.ITNBR,T1.HOUSE,T1.ITCLS,T1.MOHTQ   
FROM AMFLIBA.ITEMBL T1, AMFLIBA.ITMRVA T2, AMFLIBA.WHSMST T3 
WHERE  T2.ITCLS = T1.ITCLS AND T2.ITNBR = T1.ITNBR AND T2.STID = T3.STID AND T1.HOUSE = T3.WHID AND ((T1.HOUSE='335') AND (T1.MOHTQ<>0))
ORDER BY T1.ITNBR) as a
LEFT JOIN (SELECT PRICE.PITEM, PRICE.PAMNT FROM AFILELIB.PRICE PRICE WHERE PRICE.PRICCD='FOBARC' ORDER BY PRICE.PITEM) as b 
ON a.ITNBR = b.PITEM 
) AS x
GROUP BY x.WEEK,x."Type"
