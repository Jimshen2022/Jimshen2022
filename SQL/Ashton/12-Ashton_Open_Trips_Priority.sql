--ASHTON OPEN ORDRE QUERY FOR VO,MAY, IT WAS CREATED ON Oct.17.2022, Ver:03
SELECT d1.HOUSE,d1.CCUSNO,d1.CUSNM,d1.CSHPNO,d1.ORDUSR,d1.ALC,d1.BDTRP#,d1.LOADDATE,SUM(TRIP_QTY) AS TRIP_QTY
FROM 
(
SELECT a1.HOUSE,a1.ORDNO,A1.ITMSQ,a1.ITNBR,a1.ITDSC,a1.ITCLS,a1.CCUSNO,a1.CSHPNO,a1.CUSNM,
to_date(char(a1.TKNDAT),'yyyymmdd') Order_Taken_Date,to_date(char(a1.FRZDAT),'yyyymmdd') Original_Request_Date, to_date(char(a1.RQSDAT),'yyyymmdd') CRD,to_date(char(a1.RQIDT),'yyyymmdd') CPD, to_date(char(a1.MFIDT),'yyyymmdd')  LoadDate,
a1.ORDUSR,a1.COQTY,a1.QTYSH,a1.QTYBO,a1.OPEN_CO_QTY,a1.ALC,
a1.Product,x1.BDTRP#,x1.BDISEQ, x1.BDITQT as Trip_Qty,
x1.BDITCT,x1.BDITWT,x1.BDREF#,x1.BHCDAT,x1.BHCTIM,x1.BHRDAT,x1.BHLDAT,x1.BHLTIM

FROM 
(
Select  t1.HOUSE,t1.ORDNO,t1.ITMSQ,t1.ITNBR,t1.ITDSC,t1.ITCLS, t1.CCUSNO,t3.CUSNM, T1.CSHPNO, T1.RQIDT,T1.MFIDT,T1.UNMSR,
(CASE 
    WHEN t1.ITCLS NOT LIKE 'Z%' THEN 'RP'
	WHEN SUBSTR(t1.ITNBR,1,4)='100-' THEN 'CG'
	WHEN SUBSTR(t1.ITNBR,1,1) in ('A','B','D','E','H','L','M','P','Q','R','T','W','Z') THEN 'CG'
	ELSE 'UPH' END) as Product,t2.TKNDAT,t2.FRZDAT,t2.RQSDAT,t2.ORDUSR,
 t1.COQTY,t1.QTYSH,t1.QTYBO, T1.COQTY-T1.QTYSH AS OPEN_CO_QTY, 
(CASE 
	WHEN t1.IAFLG=0 THEN 'N' 
	WHEN t1.IAFLG = 2 THEN 'Y'
	ELSE 'Check' END) AS ALC 

FROM AFILELIB.CODATAN t1, AFILELIB.EXTORD t2,AFILELIB.ACUSMASJ t3, AFILELIB.COMAST t4, AMFLIBA.ITMRVA t5
WHERE t2.XORDNO =t1.ORDNO AND t3.CUSNO = t1.CCUSNO AND t1.ORDNO=t4.ORDNO AND t1.ITNBR = T5.ITNBR AND t1.house = T5.STID AND t1.house IN ('335')
AND t1.COQTY-t1.QTYSH<>0

) as a1

LEFT JOIN 
(-- trip demand
SELECT  t1.BDTRP#,t1.BDORD#,t1.BDISEQ,t1.BDITM#,t1.BDITMD,t1.BDCUS#, t1.BDITQT,
t1.BDITCT,t1.BDITWT,t1.BDREF#,t1.BDCDAT,t1.BDCTIM,t2.BHTRPS,t2.BHCDAT,t2.BHCTIM,t2.BHRDAT,t2.BHLDAT,t2.BHLTIM
FROM DISTLIB.BTTRIPD t1, DISTLIB.BTTRIPH t2 
WHERE t2.BHWHS# IN ('335') AND t2.BHLDAT BETWEEN 0 AND 29991231 AND t2.BHTRPS IN ('A','R','X') AND t1.BDTRP# = t2.BHTRP# 
ORDER BY t1.BDTRP#,t1.BDISEQ,t1.BDITM#
) x1  ON a1.ORDNO||a1.ITMSQ||a1.ITNBR||a1.CCUSNO = x1.BDORD#||x1.BDISEQ||x1.BDITM#||x1.BDCUS#
ORDER BY a1.MFIDT,x1.BDTRP#,a1.ITNBR,x1.BDISEQ
) AS d1
Where d1.BDTRP# IS NOT NULL
GROUP BY d1.HOUSE,d1.CCUSNO,d1.CUSNM,d1.CSHPNO,d1.ORDUSR,d1.ALC,d1.BDTRP#,d1.LOADDATE