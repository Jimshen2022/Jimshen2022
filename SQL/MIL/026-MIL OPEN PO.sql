

-- ITEM TIHI
SELECT t4.ITNBR, t4.ITNBR, t4.STID, t2.MOHTQ, t2.WHSLC, t4.ITCLS, t2.QTSYR, t4.B2Z95S, t4.ITDSC, 
t1.TIHIUNLD, t1.PICKPUT, t1.ITMCLSID, t1.UNITSWIDE, t1.UNITLAYERS, t1.UNITSDEEP, t1.SCOOPQTY, t1.SKIDSIZE, 
t3.QTYCR, t3.NBSEAT, t3.CRTWIN, t3.CRTLIN, t3.CRTHIN, t3.PRDWIN, t3.PRDHIN, t3.PRDLIN, t3.ITMWEGHT

FROM AFILELIBL.ITBEXT t1, AMFLIBL.ITEMBL t2, AFILELIBL.ITMEXT t3, AMFLIBL.ITMRVA t4
WHERE t1.ITNBR = t4.ITNBR AND t2.HOUSE = t1.HOUSE AND t2.ITNBR = t1.ITNBR AND t2.ITNBR = t4.ITNBR 
AND t2.ITCLS = t4.ITCLS AND t3.ITNBR = t1.ITNBR AND t3.ITNBR = t2.ITNBR AND t3.ITNBR = t4.ITNBR AND t4.STID = t1.HOUSE AND t4.STID IN ('51')
AND SUBSTR(t4.ITNBR,1,1) IN ('E','W') AND t4.ITCLS LIKE 'Z%' AND t4.ITCLS NOT LIKE '%K'




--ITEM OPENCO
SELECT a1.CDA3CD as WH,a1.CDAITX as ITEM,a1.CDGLCD as ITCLS, a1.CDD0NB as ETD, a1.CDB9CD as Destination,
 Date('20'||Substr(a1.CDD0NB, 2, 2) || '-'||  Substr(a1.CDD0NB, 4, 2)|| '-' ||substr(a1.CDD0NB, 6, 2)) AS "DATE",
WEEK(DATE('20'||SUBSTR(CHAR(a1.CDD0NB),2,2)||'-'||SUBSTR(CHAR(a1.CDD0NB),4,2)||'-'||SUBSTR(CHAR(a1.CDD0NB),6,2))) AS "WEEK",
YEAR(DATE('20'||SUBSTR(CHAR(a1.CDD0NB),2,2)||'-'||SUBSTR(CHAR(a1.CDD0NB),4,2)||'-'||SUBSTR(CHAR(a1.CDD0NB),6,2))) AS "YEAR"
,a1.CDAGNV as OPENCO

FROM AMFLIBL.MBCDRESM a1
WHERE CDAGNV >0 and CDGLCD like 'Z%'




-- OPEN PO

SELECT t3.ITNBR, t3.QTYOR, t4.HOUSE, t4.ORDNO, t4.VNDNR, t4.PSTTS, t3.STKQT, t3.STKDT, t3.DYLDE, t3.DYLDL, t2.UUCCIM, t3.DUEDT, t1.ITMCLSID, t1.PICKPUT, 
t1.SCOOPQTY, t1.SKIDSIZE, t3.DOKDT, t3.MSNDD, t3.MSNSD, t5.VNNMVM, t3.QTDEV, t6.ITCLS,
WEEK(DATE('20'||SUBSTR(CHAR(t3.DUEDT),2,2)||'-'||SUBSTR(CHAR(t3.DUEDT),4,2)||'-'||SUBSTR(CHAR(t3.DUEDT),6,2))) AS "WEEK",
YEAR(DATE('20'||SUBSTR(CHAR(t3.DUEDT),2,2)||'-'||SUBSTR(CHAR(t3.DUEDT),4,2)||'-'||SUBSTR(CHAR(t3.DUEDT),6,2))) AS "YEAR"

FROM AFILELIBL.ITBEXT t1, AFILELIBL.ITMEXT t2, AMFLIBL.POITEM t3, AMFLIBL.POMAST t4, AMFLIBL.VENNAML0 t5,AMFLIBL.ITMRVA t6
WHERE t1.ITNBR = t2.ITNBR AND t3.ITNBR = t2.ITNBR AND t3.ITNBR=t6.ITNBR AND t3.ORDNO = t4.ORDNO AND t4.VNDNR = t5.VNDRVM AND t3.HOUSE = t1.HOUSE AND t4.HOUSE = t3.HOUSE 
AND t4.HOUSE='51' AND t4.PSTTS IN ('10','20') AND t6.ITCLS IN ('RFR','RFRU','RFRP')  AND t3.DUEDT > '1220101'






