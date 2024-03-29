-- wvog Tihi SETUP, to pull out item class begin with Z%, create by Jimshen on Oct.08.2021

Select s1.ITNBR, s1.HOUSE, S1.MOHTQ, S1.WHSLC, S1.ITCLS, s1.QTSYR, s1.B2Z95S, 
s1.ITDSC, s1.TIHIUNLD, s1.PICKPUT, s1.ITMCLSID, s1.UNITSWIDE, s1.UNITLAYERS, s1.UNITSDEEP, 
s1.SCOOPQTY, s1.SKIDSIZE, s1.QTYCR, s1.NBSEAT, s1.CRTWIN, s1.CRTLIN, s1.CRTHIN, s1.PRDWIN, 
s1.PRDHIN, s1.PRDLIN, s1.ITMWEGHT,s2.OPEN_PO

from (SELECT t4.ITNBR, t2.HOUSE, t2.MOHTQ, t2.WHSLC, t2.ITCLS, t2.QTSYR, t4.B2Z95S, 
t4.ITDSC, t1.TIHIUNLD, t1.PICKPUT, t1.ITMCLSID, t1.UNITSWIDE, t1.UNITLAYERS, t1.UNITSDEEP, 
t1.SCOOPQTY, t1.SKIDSIZE, t3.QTYCR, t3.NBSEAT, t3.CRTWIN, t3.CRTLIN, t3.CRTHIN, t3.PRDWIN, 
t3.PRDHIN, t3.PRDLIN, t3.ITMWEGHT
FROM AFILELIBQ.ITBEXT t1, AMFLIBQ.ITEMBL t2, AFILELIBQ.ITMEXT t3, AMFLIBQ.ITMRVA t4
WHERE t1.ITNBR = t4.ITNBR AND t2.HOUSE = t1.HOUSE AND t2.ITNBR = t1.ITNBR AND 
t2.ITNBR = t4.ITNBR AND t2.ITCLS = t4.ITCLS AND t3.ITNBR = t1.ITNBR AND t3.ITNBR = t2.ITNBR AND 
t3.ITNBR = t4.ITNBR AND t4.STID = t1.HOUSE AND t1.HOUSE='232') as s1


LEFT JOIN

(SELECT c.ITNBR, SUM(c.QTYOR) AS OPEN_PO
FROM AFILELIBQ.ITBEXT a, AFILELIBQ.ITMEXT b, AMFLIBQ.POITEM c, AMFLIBQ.POMAST d, AMFLIBQ.VENNAML0 e
WHERE a.ITNBR = b.ITNBR AND c.ITNBR = b.ITNBR AND c.ORDNO = d.ORDNO AND 
d.VNDNR = e.VNDRVM AND c.HOUSE = a.HOUSE AND d.HOUSE = c.HOUSE
AND d.HOUSE='232' AND d.PSTTS NOT IN ('10','20','30') and c.DUEDT  
BETWEEN  int('1'||substr(trim(char(CURRENT DATE - 30 days)),3,2)||substr(trim(char(CURRENT DATE- 30 days)),6,2)||substr(trim(char(CURRENT DATE- 30 days)),9,2)) 
AND int('1'||substr(trim(char(CURRENT DATE + 360 days)),3,2)||substr(trim(char(CURRENT DATE + 360 days)),6,2)||substr(trim(char(CURRENT DATE + 360 days)),9,2))
GROUP BY C.ITNBR) as s2

on s1.ITNBR = s2.ITNBR

WHERE  s1.ITCLS LIKE 'Z%' and s1.ITCLS not LIKE '%K'