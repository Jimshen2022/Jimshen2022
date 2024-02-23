SELECT t4.ITNBR, t2.ITNBR, t2.HOUSE, t2.MOHTQ, t2.WHSLC, t2.ITCLS, t2.QTSYR, t4.B2Z95S, t4.ITDSC, t1.TIHIUNLD, t1.PICKPUT, t1.ITMCLSID, t1.UNITSWIDE, t1.UNITLAYERS, t1.UNITSDEEP, t1.SCOOPQTY, t1.SKIDSIZE, t3.QTYCR, t3.NBSEAT, t3.CRTWIN, t3.CRTLIN, t3.CRTHIN, t3.PRDWIN, t3.PRDHIN, t3.PRDLIN, t3.ITMWEGHT

FROM AFILELIBL.ITBEXT t1, AMFLIBL.ITEMBL t2, AFILELIBL.ITMEXT t3, AMFLIBL.ITMRVA t4

WHERE t1.ITNBR = t4.ITNBR AND t2.HOUSE = t1.HOUSE AND t2.ITNBR = t1.ITNBR AND t2.ITNBR = t4.ITNBR AND t2.ITCLS = t4.ITCLS AND t3.ITNBR = t1.ITNBR AND t3.ITNBR = t2.ITNBR AND t3.ITNBR = t4.ITNBR AND t4.STID = t1.HOUSE AND (t1.HOUSE='51') AND (t2.ITCLS like 'Z%') and (t2.ITCLS not like '%K')
