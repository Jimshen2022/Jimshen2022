
-- ON HAND IN WHSE232 ON DSN=ASHLEYSHBIB
SELECT x.ITNBR,x.ITCLS, x.Product, x."UP($USD)",SUM(x.MOHTQ) AS MOHTQ, SUM(x."AMT($USD)") AS "AMT($USD)"
FROM
(
SELECT a.ITNBR,a.HOUSE, a.ITCLS, a.MOHTQ, b.PAMNT "UP($USD)",
(CASE 
    WHEN a.ITCLS NOT LIKE 'Z%' THEN 'RP'
	WHEN SUBSTR(a.ITNBR,1,4)='100-' THEN 'CG'
	WHEN SUBSTR(a.ITNBR,1,5)='5100-' THEN 'CG'	
	WHEN SUBSTR(a.ITNBR,1,1) in ('A','L','Q','R') THEN 'ACCESSORIES'
	WHEN SUBSTR(a.ITNBR,1,1) in ('B','D','E','H','P','T','W','Z') THEN 'CG'
	ELSE 'UPH' END) as Product,
(CASE 
    WHEN b.PAMNT IS NULL AND a.ITCLS NOT LIKE 'Z%' THEN 0
	WHEN b.PAMNT IS NULL AND a.ITCLS LIKE 'Z%' THEN a.MOHTQ*50
	ELSE a.MOHTQ*b.PAMNT END) as "AMT($USD)"
FROM 
(SELECT T1.ITNBR,T1.HOUSE, T1.ITCLS, T1.MOHTQ, T1.WHSLC, T2.ITDSC    
FROM AMFLIBQ.ITEMBL T1, AMFLIBQ.ITMRVA T2, AMFLIBQ.WHSMST T3 
WHERE  T2.ITCLS = T1.ITCLS AND T2.ITNBR = T1.ITNBR AND T2.STID = T3.STID AND T1.HOUSE = T3.WHID AND (T1.HOUSE IN ('232','21','PC1','22') AND T1.MOHTQ<>0)
ORDER BY T1.ITNBR) as a
LEFT JOIN (SELECT PRICE.PITEM, MAX(PRICE.PAMNT)/6.37 as PAMNT FROM AFILELIBQ.PRICE PRICE GROUP BY PRICE.PITEM ORDER BY PRICE.PITEM) as b 
ON a.ITNBR = b.PITEM 
) x
GROUP BY x.ITNBR,x.ITCLS, x.Product, x."UP($USD)"
ORDER BY x.ITNBR



-- WH Trx Summary in qty and amount  DSN=ASHLEYSHBIA

SELECT X.ITNBR, X.HOUSE, X.ITCLS,X.WEEK,X.YEAR,X.YEAR*100+X.WEEK AS "YearWeek", X.TRQTY,Y.PAMNT "UP($USD)",
(CASE 
    WHEN X.ITCLS NOT LIKE 'Z%' THEN 'RP'
	WHEN SUBSTR(X.ITNBR,1,4)='100-' THEN 'CG'
	WHEN SUBSTR(X.ITNBR,1,5)='5100-' THEN 'CG'
	WHEN SUBSTR(X.ITNBR,1,1) in ('A','L','Q','R') THEN 'ACCESSORIES'
	WHEN SUBSTR(X.ITNBR,1,1) in ('B','D','E','H','P','T','W','Z') THEN 'CG'
	ELSE 'UPH' END) as Product,
(CASE 
		WHEN X.TCODE IN ('RP') THEN 'IN'
		WHEN X.TCODE IN ('SA') THEN 'OUT'
		ELSE 'CHECK' END ) as "TYPE",
(CASE 
    WHEN Y.PAMNT IS NULL AND X.ITCLS NOT LIKE 'Z%' THEN 0
	WHEN Y.PAMNT IS NULL AND X.ITCLS LIKE 'Z%' THEN X.TRQTY*50
	ELSE X.TRQTY*Y.PAMNT END) as "AMT($USD)"	
FROM
(
SELECT T1.TCODE, T1.ITNBR, T1.HOUSE, T2.ITCLS, WEEK(DATE('20'||SUBSTR(CHAR(T1.UPDDT),2,2)||'-'||SUBSTR(CHAR(T1.UPDDT),4,2)||'-'||SUBSTR(CHAR(T1.UPDDT),6,2))) AS WEEK,
YEAR(DATE('20'||SUBSTR(CHAR(T1.UPDDT),2,2)||'-'||SUBSTR(CHAR(T1.UPDDT),4,2)||'-'||SUBSTR(CHAR(T1.UPDDT),6,2))) AS YEAR,
SUM(T1.TRQTY) TRQTY
FROM AMFLIBQ.IMHIST T1, AMFLIBQ.ITMRVA T2, AMFLIBQ.WHSMST T3
WHERE T2.ITNBR = T1.ITNBR AND T2.STID = T3.STID AND T1.HOUSE = T3.WHID AND T1.HOUSE IN ('232','21','PC1') 
AND T1.UPDDT BETWEEN CHAR('1'||VARCHAR_FORMAT(current date - 40 Days,'YYMMDD')) AND CHAR('1'||VARCHAR_FORMAT(current date,'YYMMDD'))
AND T1.TRQTY<>0 AND T1.TCODE IN ('RP','SA') 
GROUP BY T1.TCODE, T1.ITNBR, T1.HOUSE, T2.ITCLS,WEEK(DATE('20'||SUBSTR(CHAR(T1.UPDDT),2,2)||'-'||SUBSTR(CHAR(T1.UPDDT),4,2)||'-'||SUBSTR(CHAR(T1.UPDDT),6,2))),
YEAR(DATE('20'||SUBSTR(CHAR(T1.UPDDT),2,2)||'-'||SUBSTR(CHAR(T1.UPDDT),4,2)||'-'||SUBSTR(CHAR(T1.UPDDT),6,2)))
ORDER BY T1.ITNBR 
) AS X 
LEFT JOIN (SELECT PRICE.PITEM, MAX(PRICE.PAMNT)/6.37 as PAMNT FROM AFILELIBQ.PRICE PRICE GROUP BY PRICE.PITEM ORDER BY PRICE.PITEM) as Y 
ON X.ITNBR = Y.PITEM




