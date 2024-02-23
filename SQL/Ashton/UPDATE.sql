UPDATE
(
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
) AS X
SET X.ITNBR = "JIMSHEN" WHERE X.ITNBR = 'A3000024'
)


