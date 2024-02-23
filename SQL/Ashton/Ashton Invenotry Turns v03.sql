
-- ON HAND IN WHSE335 ON DSN=AFIPROD
SELECT a.ITNBR,a.HOUSE, a.ITCLS, a.MOHTQ, b.PAMNT "UP($USD)",
(CASE 
    WHEN a.ITCLS NOT LIKE 'Z%' THEN 'RP'
	WHEN SUBSTR(a.ITNBR,1,4)='100-' THEN 'CG'
	WHEN SUBSTR(a.ITNBR,1,1) in ('A','B','D','E','H','L','P','Q','M','R','T','W','Z') THEN 'CG'
	ELSE 'UPH' END) as Product,
(CASE 
    WHEN b.PAMNT IS NULL AND a.ITCLS NOT LIKE 'Z%' THEN 0
	WHEN b.PAMNT IS NULL AND a.ITCLS LIKE 'Z%' THEN a.MOHTQ*50
	ELSE a.MOHTQ*b.PAMNT END) as "AMT($USD)"
FROM 
(SELECT T1.ITNBR,T1.HOUSE, T1.ITCLS, T1.MOHTQ, T1.WHSLC, T2.ITDSC    
FROM AMFLIBA.ITEMBL T1, AMFLIBA.ITMRVA T2, AMFLIBA.WHSMST T3 
WHERE  T2.ITCLS = T1.ITCLS AND T2.ITNBR = T1.ITNBR AND T2.STID = T3.STID AND T1.HOUSE = T3.WHID AND ((T1.HOUSE='335') AND (T1.MOHTQ<>0))
ORDER BY T1.ITNBR) as a
LEFT JOIN (SELECT PRICE.PITEM, PRICE.PAMNT FROM AFILELIB.PRICE PRICE WHERE PRICE.PRICCD='FOBARC' ORDER BY PRICE.PITEM) as b 
ON a.ITNBR = b.PITEM 




-- WH Trx Summary in qty and amount  DSN=AFIBI
SELECT X.ITNBR, X.HOUSE, X.ITCLS,X.WEEK,X.YEAR,X.YEAR*100+X.WEEK AS "YearWeek",X.TRQTY,Y.PAMNT "UP($USD)",
(CASE 
    WHEN X.ITCLS NOT LIKE 'Z%' THEN 'RP'
	WHEN SUBSTR(X.ITNBR,1,4)='100-' THEN 'CG'
	WHEN SUBSTR(X.ITNBR,1,1) in ('A','B','D','E','H','L','P','Q','M','R','T','W','Z') THEN 'CG'
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
FROM AMFLIBA.IMHIST T1, AMFLIBA.ITMRVA T2, AMFLIBA.WHSMST T3
WHERE T2.ITNBR = T1.ITNBR AND T2.STID = T3.STID AND T1.HOUSE = T3.WHID AND T1.HOUSE='335' 
AND T1.UPDDT BETWEEN CHAR('1'||VARCHAR_FORMAT(current date - 30 Days,'YYMMDD')) AND CHAR('1'||VARCHAR_FORMAT(current date,'YYMMDD'))
AND T1.TRQTY<>0 AND T1.TCODE IN ('RP','SA') 
GROUP BY T1.TCODE, T1.ITNBR, T1.HOUSE, T2.ITCLS,WEEK(DATE('20'||SUBSTR(CHAR(T1.UPDDT),2,2)||'-'||SUBSTR(CHAR(T1.UPDDT),4,2)||'-'||SUBSTR(CHAR(T1.UPDDT),6,2))),
YEAR(DATE('20'||SUBSTR(CHAR(T1.UPDDT),2,2)||'-'||SUBSTR(CHAR(T1.UPDDT),4,2)||'-'||SUBSTR(CHAR(T1.UPDDT),6,2)))
ORDER BY T1.ITNBR 
) AS X 
LEFT JOIN (SELECT PRICE.PITEM, PRICE.PAMNT FROM AFILELIB.PRICE PRICE WHERE PRICE.PRICCD='FOBARC' ORDER BY PRICE.PITEM) as Y 
ON X.ITNBR = Y.PITEM











