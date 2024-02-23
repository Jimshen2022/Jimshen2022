-- WH232-22 ON HAND LIST

SELECT T1.ITNBR, T1.HOUSE, T1.ITCLS, T2.ITDSC, T4.LLOCN, T4.LQNTY
FROM AMFLIBQ.ITEMBL T1, AMFLIBQ.ITMRVA T2, AMFLIBQ.WHSMST T3, AMFLIBQ.SLQNTY T4
WHERE T1.ITNBR = T4.ITNBR AND T4.ITNBR = T2.ITNBR AND T1.ITCLS= T2.ITCLS AND T1.HOUSE=T4.HOUSE AND T4.HOUSE = T3.WHID AND T3.STID = T2.STID AND ((T1.MOHTQ<>0) AND (T1.HOUSE in ('22','232','21')))
ORDER BY T1.ITNBR


-- WH232,22 ON HAND LIST V02
SELECT T1.ITNBR, T1.ITCLS, T2.ITDSC, 
(CASE 
    WHEN t1.ITCLS NOT LIKE 'Z%' THEN 'RP'
	WHEN SUBSTR(t1.ITNBR,1,4)='100-' THEN 'CG'
	WHEN SUBSTR(t1.ITNBR,1,1) in ('A','B','D','H','L','P','Q','R','T','W','Z') THEN 'CG'
	ELSE 'UPH' END) as Product,
SUM(CASE WHEN T1.HOUSE IN ('232') THEN T4.LQNTY ELSE 0 END) AS WH232,
SUM(CASE WHEN T1.HOUSE IN ('22') THEN T4.LQNTY ELSE 0 END) AS WH22,
SUM(CASE WHEN T1.HOUSE IN ('232') THEN T4.LQNTY ELSE 0 END)+SUM(CASE WHEN T1.HOUSE IN ('22') THEN T4.LQNTY ELSE 0 END) AS TotalOnHand
FROM AMFLIBQ.ITEMBL T1, AMFLIBQ.ITMRVA T2, AMFLIBQ.WHSMST T3, AMFLIBQ.SLQNTY T4
WHERE T1.ITNBR = T4.ITNBR AND T4.ITNBR = T2.ITNBR AND T1.ITCLS= T2.ITCLS AND T1.HOUSE=T4.HOUSE AND T4.HOUSE = T3.WHID AND T3.STID = T2.STID AND ((T1.MOHTQ<>0) AND (T1.HOUSE in ('22','232','21')))
GROUP BY T1.ITNBR, T1.ITCLS, T2.ITDSC,(CASE 
    WHEN t1.ITCLS NOT LIKE 'Z%' THEN 'RP'
	WHEN SUBSTR(t1.ITNBR,1,4)='100-' THEN 'CG'
	WHEN SUBSTR(t1.ITNBR,1,1) in ('A','B','D','H','L','P','Q','R','T','W','Z') THEN 'CG'
	ELSE 'UPH' END)
ORDER BY T1.ITNBR












-- AS ON HAND

SELECT t1.ITNBR,t1.ITCLS, SUM(t1.MOHTQ) as OnHand,
(CASE 
    WHEN t1.ITCLS NOT LIKE 'Z%' THEN 'RP'
	WHEN SUBSTR(t1.ITNBR,1,4)='100-' THEN 'CG'
	WHEN SUBSTR(t1.ITNBR,1,1) in ('A','B','D','H','L','P','Q','R','T','W','Z') THEN 'CG'
	ELSE 'UPH' END) as Product
	
FROM AMFLIBQ.ITEMBL t1, AMFLIBQ.ITMRVA t2, AMFLIBQ.WHSMST t3
WHERE t2.ITCLS = t1.ITCLS AND t2.ITNBR = t1.ITNBR AND t1.HOUSE = t3.WHID 
AND t3.STID = t2.STID AND ((t1.MOHTQ<>0) 
AND (t1.HOUSE='232' Or t1.HOUSE='21'))
GROUP BY t1.ITNBR,t1.ITCLS
ORDER BY t1.ITNBR








-- AS ONHAND SUMMARY

SELECT z1.PRODUCT,z1.ITNBR,Sum(z1.OnHand) as OH,'OnHand' as "TYPE"
from 
(
SELECT t1.ITNBR,t1.ITCLS, SUM(t1.MOHTQ) as OnHand,
(CASE 
    WHEN t1.ITCLS NOT LIKE 'Z%' THEN 'RP'
	WHEN SUBSTR(t1.ITNBR,1,4)='100-' THEN 'CG'
	WHEN SUBSTR(t1.ITNBR,1,1) in ('A','B','D','H','L','P','Q','R','T','W','Z') THEN 'CG'
	ELSE 'UPH' END) as Product
	
FROM AMFLIBQ.ITEMBL t1, AMFLIBQ.ITMRVA t2, AMFLIBQ.WHSMST t3
WHERE t2.ITCLS = t1.ITCLS AND t2.ITNBR = t1.ITNBR AND t1.HOUSE = t3.WHID 
AND t3.STID = t2.STID AND ((t1.MOHTQ<>0) 
AND (t1.HOUSE='232' Or t1.HOUSE='21'))
GROUP BY t1.ITNBR,t1.ITCLS
ORDER BY t1.ITNBR
)  as z1
group by z1.PRODUCT,z1.ITNBR

