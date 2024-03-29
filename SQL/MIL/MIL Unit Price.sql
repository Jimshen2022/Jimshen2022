-- MIL UNIT PRICE CREATED ON Feb.15.2022 BY JIMSHEN

SELECT x.RPAITX, x.ITCLS, x.RPAMVA, x.RPBLDT, x.RPZ0D7

FROM
(
((SELECT a.RPAITX,(CASE WHEN a.RPBRCD IN ('VND') THEN a.RPAMVA/22715 ELSE a.RPAMVA END) AS RPAMVA,a.RPBLDT,a.RPZ0D7, T2.ITCLS
FROM AMFLIBL.ITMFPR a 
LEFT JOIN AMFLIBL.ITMRVA T2 ON a.RPAITX=T2.ITNBR AND a.RPZ0D7 = T2.STID 
WHERE a.RPZ0D7 = '51' AND a.RPAITX||a.RPZ0D7||a.RPBLDT IN (SELECT a.RPAITX||a.RPZ0D7||MAX(a.RPBLDT) RPBLDT FROM AMFLIBL.ITMFPR a  WHERE a.RPZ0D7 = '51' GROUP BY a.RPAITX,a.RPZ0D7)) 
UNION ALL
(SELECT t1.ITNO1G, t1.UCCT1G/22715 AS RPAMVA, t1.CCDT1G, t1.STID1G, t1.STID1G FROM AMFLIBL.ITMPRB t1))
UNION ALL
(SELECT t1.ITNBR, t1.LCOST/22715 AS RPAMVA, t1.LDQOH, t1.HOUSE, t1.ITCLS FROM AMFLIBL.ITEMBL t1))AS x
ORDER BY x.RPAITX, x.RPAMVA ASC

