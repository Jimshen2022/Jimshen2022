SELECT Y1.RPLCONTAINERNUMBER, Y1.WCSORIGIN, Y1.WCSDESTINATION, Y1.WCSORDER, Y1.RPLITEMNUMBER, Y1.ITCLS, Y1.DATE, Y1.HOUR, 
Y1.RPLLOADUSERNAME,Y1.ITMCQTY, Y1.B2Z95S, Y1.Product, Y1.LOADED_QTY, Y1.BoxQty, Y1.CUBES

FROM 
(SELECT t1.RPLCONTAINERNUMBER, '51' AS WCSORIGIN, 'Null' AS WCSDESTINATION, 'Null' as WCSORDER, t1.RPLITEMNUMBER, t3.ITCLS, DATE(t1.RPLLOADDATE) AS DATE, HOUR(t1.RPLLOADDATE) AS HOUR, 
t1.RPLLOADUSERNAME,t2.ITMCQTY, t3.B2Z95S, 
(CASE 
        WHEN t3.ITCLS IN ('WPLS') THEN 'Plastics'
		WHEN t3.ITCLS IN ('PANL') THEN 'Panel'
        WHEN t3.ITCLS IN ('WVBC','WVCS') THEN 'Foundation'
        WHEN t3.ITCLS IN ('SLDK') THEN 'RP'
        WHEN t3.ITCLS LIKE 'T%' THEN 'RP'
		WHEN t3.ITCLS LIKE 'R%' THEN 'RP'
        WHEN t3.ITCLS IN ('ZKIS') THEN 'Bedding'
		WHEN t3.ITCLS IN ('ZKIZ') THEN 'ZipperCover'
        WHEN t3.ITCLS LIKE 'Z%' AND t3.ITCLS LIKE '%K' THEN 'UnKits'
		WHEN t3.ITCLS IN ('PACS') THEN 'UnKits'
        WHEN t3.ITCLS IN ('BBFR','WVHC') THEN 'Verona'
        WHEN t3.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB') THEN 'CG'
        WHEN t3.ITCLS IN ('ZACM','ZASU','ZMLH','ZMLR','ZUSR','ZUSU','ZVUC','ZXUC') THEN 'UPH'
        ELSE 'RawMaterial' END) AS Product, 
SUM(t1.RPLPRINTQUANTITY) AS LOADED_QTY, SUM(CAST(t1.RPLPRINTQUANTITY AS FLOAT))/t2.ITMCQTY AS BoxQty, SUM(t1.RPLPRINTQUANTITY)*t3.B2Z95S AS CUBES

FROM RGNFILL.PC228RPF t1, AFILELIBL.ITMEXT t2, AMFLIBL.ITMRVA t3
WHERE t1.RPLLOADDATE BETWEEN ? AND ? AND t1.RPLITEMNUMBER = t2.ITNBR AND t1.RPLITEMNUMBER = t3.itnbr
GROUP BY 
t1.RPLCONTAINERNUMBER,'51', 'Null', 'Null', t1.RPLITEMNUMBER, t3.ITCLS, DATE(t1.RPLLOADDATE), HOUR(t1.RPLLOADDATE), 
t1.RPLLOADUSERNAME,t2.ITMCQTY, t3.B2Z95S,
(CASE 
        WHEN t3.ITCLS IN ('WPLS') THEN 'Plastics'
		WHEN t3.ITCLS IN ('PANL') THEN 'Panel'
        WHEN t3.ITCLS IN ('WVBC','WVCS') THEN 'Foundation'
        WHEN t3.ITCLS IN ('SLDK') THEN 'RP'
        WHEN t3.ITCLS LIKE 'T%' THEN 'RP'
		WHEN t3.ITCLS LIKE 'R%' THEN 'RP'
        WHEN t3.ITCLS IN ('ZKIS') THEN 'Bedding'
		WHEN t3.ITCLS IN ('ZKIZ') THEN 'ZipperCover'
        WHEN t3.ITCLS LIKE 'Z%' AND t3.ITCLS LIKE '%K' THEN 'UnKits'
		WHEN t3.ITCLS IN ('PACS') THEN 'UnKits'
        WHEN t3.ITCLS IN ('BBFR','WVHC') THEN 'Verona'
        WHEN t3.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB') THEN 'CG'
        WHEN t3.ITCLS IN ('ZACM','ZASU','ZMLH','ZMLR','ZUSR','ZUSU','ZVUC','ZXUC') THEN 'UPH'
        ELSE 'RawMaterial' END)
) AS Y1

RIGHT JOIN

(SELECT 
a.WCHCONTAINERNUMBER,a.WCHORIGIN,a.WCHDESTINATION,a.WCHCONTAINERSTATUS,a.WCHTOTALCARTONS,a.WCHTOTALCUBES,a.WCHPOSTEDTIMESTAMP,a.WCHTOTALWEIGHT,
a.WCHCONTAINERSIZE,
trim(a.WCHORIGIN)||'-'|| trim(a.WCHCONTAINERNUMBER)||'-'||trim(a.WCHDESTINATION) as Container#
FROM  LLUSAF.WVCNTHD a
WHERE a.WCHCONTAINERSTATUS not in ('P','T') AND a.WCHORIGIN IN ('51')  AND a.WCHPOSTEDTIMESTAMP BETWEEN char(current date - 30 days) and char(current DATE) 
) AS Y2 ON Y1.RPLCONTAINERNUMBER = Y2.WCHCONTAINERNUMBER
