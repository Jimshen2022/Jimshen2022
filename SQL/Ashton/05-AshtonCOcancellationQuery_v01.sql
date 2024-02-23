
SELECT t2.ITNBR,t2.ITCLS,t2.HOUSE,t2.ORDNO,t2.COQTY,t2.QTYSH,t2.QTYBO,t2.UNMSR,
t2."PromiseDate",t2."LoadDate",t2.CCUSNO,x1."OrderTakenDate",x1."CustReqstDate",x1."UserLastMaintain",x1."DateLastMaintain",
x1.TIMMNT,x1."FrozenCustRequestDate",x1.CUSTNO  
FROM 
(SELECT t1.ITNBR,t1.ITCLS,t1.HOUSE,t1.ORDNO,t1.COQTY,t1.QTYSH,t1.QTYBO,t1.UNMSR,
t1.RQIDT AS "PromiseDate",t1.MFIDT AS "LoadDate",t1.CCUSNO
FROM AFILELIB.CODATAN t1
WHERE t1.HOUSE IN ('335') 
AND EXISTS (SELECT 1 FROM AFILELIB.COMAST a1 WHERE t1.ORDNO = a1.ORDNO)
) t2

LEFT JOIN 

(SELECT b1.WHSE,b1.XORDNO,b1.TKNDAT as "OrderTakenDate",b1.RQSDAT as "CustReqstDate",b1.USRMNT as "UserLastMaintain",
b1.DATMNT as "DateLastMaintain",b1.TIMMNT,b1.FRZDAT as "FrozenCustRequestDate",b1.CUSTNO  
FROM AFILELIB.EXTORD b1) x1 

ON t2.ORDNO=x1.XORDNO and t2.HOUSE=x1.WHSE
