-- MIL ON HAND CREATED BY JIMSHEN UPDATED ON MAY.28.2022
SELECT a1.ITNBR, a1.ITDSC,a1.B2Z95S*a1.LQNTY as Cubes, a1.ITCLS, a1.HOUSE, a1.LLOCN, a1.LQNTY, a1.ORDNO, a1.LBHNO, a2.RPAMVA AS "UnitPrice($USD)", a1.LQNTY*a2.RPAMVA AS "AMT($USD)",
(CASE 
        WHEN a1.ITCLS IN ('SLDK') THEN 'RP'
        WHEN a1.ITCLS LIKE 'T%' THEN 'RP'
		WHEN a1.ITCLS LIKE 'R%' THEN 'RP'
		WHEN a1.ITCLS IN ('PACS') THEN 'UnKits'
        WHEN a1.ITCLS LIKE 'Z%' AND a1.ITCLS LIKE '%K' THEN 'UnKits'
        WHEN a1.ITCLS IN ('ZACM','ZASU','ZMLH','ZMLR','ZUSR','ZUSU','ZVUC','ZXUC','ZUSU','ZUMU') THEN 'UPH'
        WHEN a1.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB') THEN 'CG'
        WHEN a1.ITCLS IN ('ZKIS') THEN 'Bedding'		
        WHEN a1.ITCLS IN ('WPLS') THEN 'Plastics'
        WHEN a1.ITCLS IN ('WVBC','WVCS') THEN 'Foundation'		
		WHEN a1.ITCLS IN ('PANL') THEN 'Panel'
		WHEN a1.ITCLS IN ('ZKIZ') THEN 'ZipperCover'
        WHEN a1.ITCLS IN ('BBFR','WVHC') THEN 'Verona'		
		WHEN a1.ITCLS NOT LIKE 'Z%' THEN 'RawMaterial'
        ELSE 'Check' END) AS Product
FROM
(
-- MIL Inv Qty
SELECT t1.ITNBR, t2.ITDSC, t2.ITCLS, t1.HOUSE, t1.LLOCN, t1.FDATE, t1.LQNTY, t1.ORDNO, t1.LBHNO, t2.B2Z95S
FROM AMFLIBL.SLQNTY t1 left join AMFLIBL.ITMRVA t2 on t1.itnbr = t2.itnbr
WHERE t1.LLOCN NOT IN ('S01ST1','S01PS1')
) a1

LEFT JOIN 
(
-- MIL UNIT PRICE CREATED ON Feb.15.2022 BY JIMSHEN
SELECT b1.RPAITX, MAX(b1.RPAMVA) as RPAMVA
FROM 
(SELECT x.RPAITX, x.ITCLS, x.RPAMVA, x.RPBLDT, x.RPZ0D7
FROM
(
((SELECT a.RPAITX,(CASE WHEN a.RPBRCD IN ('VND') THEN a.RPAMVA/23090 ELSE a.RPAMVA END) AS RPAMVA,a.RPBLDT,a.RPZ0D7, T2.ITCLS
FROM AMFLIBL.ITMFPR a 
LEFT JOIN AMFLIBL.ITMRVA T2 ON a.RPAITX=T2.ITNBR AND a.RPZ0D7 = T2.STID 
WHERE a.RPZ0D7 = '51' AND a.RPAITX||a.RPZ0D7||a.RPBLDT IN (SELECT a.RPAITX||a.RPZ0D7||MAX(a.RPBLDT) RPBLDT FROM AMFLIBL.ITMFPR a  WHERE a.RPZ0D7 = '51' GROUP BY a.RPAITX,a.RPZ0D7)) 
UNION ALL
(SELECT t1.ITNO1G, t1.UCCT1G/23090 AS RPAMVA, t1.CCDT1G, t1.STID1G, t1.STID1G FROM AMFLIBL.ITMPRB t1))
UNION ALL
(SELECT t1.ITNBR, t1.LCOST/23090 AS RPAMVA, t1.LDQOH, t1.HOUSE, t1.ITCLS FROM AMFLIBL.ITEMBL t1))AS x
ORDER BY x.RPAITX, x.RPAMVA ASC) b1
GROUP BY b1.RPAITX
) a2  
ON a1.ITNBR = a2.RPAITX















-- MIL AS400 STOCK - 20210813

SELECT t2.ITNBR, t1.ITDSC, t1.ITCLS, t2.HOUSE, t2.LLOCN, t2.FDATE, t2.LQNTY, t2.LBHNO,t3.ITMCQTY,
(CASE 
        WHEN t1.ITCLS IN ('WPLS') THEN 'Plastics'
		WHEN t1.ITCLS IN ('PANL') THEN 'Panel'
        WHEN t1.ITCLS IN ('WVBC','WVCS') THEN 'Foundation'
        WHEN t1.ITCLS IN ('SLDK') THEN 'RP'
        WHEN t1.ITCLS LIKE 'T%' THEN 'RP'
		WHEN t1.ITCLS LIKE 'R%' THEN 'RP'
        WHEN t1.ITCLS IN ('ZKIS') THEN 'Bedding'
		WHEN t1.ITCLS IN ('ZKIZ') THEN 'ZipperCover'
        WHEN t1.ITCLS LIKE 'Z%' AND t1.ITCLS LIKE '%K' THEN 'UnKits'
		WHEN t1.ITCLS IN ('PACS') THEN 'UnKits'
        WHEN t1.ITCLS IN ('BBFR','WVHC') THEN 'Verona'
        WHEN t1.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB') THEN 'CG'
        WHEN t1.ITCLS IN ('ZACM','ZASU','ZMLH','ZMLR','ZUSR','ZUSU','ZVUC','ZXUC') THEN 'UPH'
        ELSE 'RawMaterial' END) AS Product
		
FROM AMFLIBL.ITEMASA t1, AMFLIBL.SLQNTY t2,AFILELIBL.ITMEXT t3
WHERE t1.ITNBR = t2.ITNBR and t2.ITNBR = t3.ITNBR AND t2.HOUSE IN ('51','52')