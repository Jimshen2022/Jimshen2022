
-- Query US inventory list
SELECT ITEMBL.ITNBR, ITEMBL.HOUSE, ITEMBL.ITCLS, ITEMBL.MOHTQ, ITEMBL.WHSLC, ITMRVA.ITDSC
FROM AMFLIBA.ITEMBL ITEMBL, AMFLIBA.ITMRVA ITMRVA, AMFLIBA.WHSMST WHSMST
WHERE ITMRVA.ITCLS = ITEMBL.ITCLS AND ITMRVA.ITNBR = ITEMBL.ITNBR AND ITMRVA.STID = WHSMST.STID 
AND ITEMBL.HOUSE = WHSMST.WHID AND ITEMBL.HOUSE in ('ECR') AND ITEMBL.MOHTQ<>0

















