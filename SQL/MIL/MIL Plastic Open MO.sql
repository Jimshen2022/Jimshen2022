-- MIL OPEN MO BY P&IC, Created by Jimshen on Sep.28.2021

SELECT * 
FROM
(
SELECT t1.STID, t1.ORDNO, t1.FITEM, t1.ORQTY+t1.QTDEV as MO_QTY, t1.QTYRC AS RECEIVED_QTY,t1.ORQTY+t1.QTDEV-t1.QTYRC AS OPEN_QTY,
t1.ODUDT,t1.ITCL, t1.JOBNO,t1.CRDT, t1.CRUS,
(CASE 
        WHEN t1.ITCL IN ('WPLS') THEN 'Plastics'
		WHEN t1.ITCL IN ('PANL') THEN 'Panel'
		WHEN t1.ITCL IN ('DECK','QA') THEN 'RawMaterial'
        WHEN t1.ITCL IN ('WVBC','WVHC','WVCS') THEN 'Foundation'
        WHEN t1.ITCL IN ('SLDK') THEN 'RP'
        WHEN t1.ITCL LIKE 'T%' THEN 'RP'
        WHEN t1.ITCL IN ('ZKIS') THEN 'Bedding'
		WHEN t1.ITCL IN ('ZKIZ') THEN 'ZipperCover'
        WHEN t1.ITCL LIKE 'Z%' AND t1.ITCL LIKE '%K' THEN 'UnKits'
		WHEN t1.ITCL IN ('PACS') THEN 'UnKits'
        WHEN t1.ITCL IN ('BBFR') THEN 'Verona'
        WHEN t1.ITCL IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB') THEN 'CG'
        WHEN t1.ITCL LIKE 'Z%' THEN 'UPH'
        ELSE 'Check' END) AS Product
		
FROM AMFLIBL.MOMAST t1
WHERE t1.ORQTY+t1.QTDEV <>0
AND t1.OSTAT in ('10','40','45') 
AND t1.ODUDT BETWEEN CHAR('1'||VARCHAR_FORMAT(current date - 90 Days,'YYMMDD')) AND CHAR('1'||VARCHAR_FORMAT(current date,'YYMMDD'))
ORDER BY t1.FITEM, t1.ODUDT, t1.ORDNO
) as x1
where x1.Product = 'Plastics'
