SELECT POITEM.ITNBR, POITEM.QTYOR, POMAST.HOUSE, POMAST.ORDNO, POMAST.VNDNR, POMAST.PSTTS, POITEM.STKQT, 
POITEM.STKDT, POITEM.DYLDE, POITEM.DYLDL, ITMEXT.UUCCIM, POITEM.DUEDT, ITBEXT.ITMCLSID, ITBEXT.PICKPUT, 
ITBEXT.SCOOPQTY, ITBEXT.SKIDSIZE, POITEM.DOKDT, POITEM.MSNDD, POITEM.MSNSD, VENNAML0.VNNMVM, POITEM.QTYOR, 
POITEM.QTDEV, POITEM.STKQT,ITEMBL.ITCLS 

FROM AFILELIB.ITBEXT ITBEXT, AFILELIB.ITMEXT ITMEXT, AMFLIBA.POITEM POITEM, AMFLIBA.POMAST POMAST, AMFLIBA.VENNAML0 VENNAML0,AMFLIBA.ITEMBL ITEMBL
WHERE ITBEXT.ITNBR = ITMEXT.ITNBR AND POITEM.ITNBR = ITMEXT.ITNBR AND ITMEXT.ITNBR=ITEMBL.ITNBR AND POITEM.ORDNO = POMAST.ORDNO 
AND POMAST.VNDNR = VENNAML0.VNDRVM AND POITEM.HOUSE = ITBEXT.HOUSE AND POMAST.HOUSE = POITEM.HOUSE 
AND POITEM.HOUSE = ITEMBL.HOUSE AND ((POMAST.PSTTS='10') OR (POMAST.PSTTS='20') OR (POMAST.PSTTS='30')) 
AND (POITEM.HOUSE='335') and (ITEMBL.ITCLS not like 'Z%') 

ORDER BY POITEM.ITNBR, POITEM.DUEDT