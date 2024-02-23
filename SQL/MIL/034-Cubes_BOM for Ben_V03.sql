-- MIL ON HAND WITH BOX_MRP AND BOX_DES CREATED BY JIMSHEN UPDATED ON Aug.24.2022
SELECT a1.ITNBR, a1.ITDSC,a1.B2Z95S*a1.MOHTQ as Cubes, a1.ITCLS, a1.HOUSE, a1.WHSLC, a1.MOHTQ, 
(CASE 
        WHEN a1.ITCLS LIKE 'TAF%' THEN 'RP'
		WHEN a1.ITCLS IN ('PACS') THEN 'UnKits'
        WHEN a1.ITCLS LIKE 'Z%' AND a1.ITCLS LIKE '%K' THEN 'UnKits'
        WHEN a1.ITCLS IN ('ZACM','ZASU','ZMLH','ZMLR','ZUSR','ZUSU','ZVUC','ZXUC','ZUSU','ZUMU','ZAMU','ZASM','ZASR','ZDMA','ZMUC','ZSUS','ZUMS','ZUSM','ZVMA','ZVUS','ZXLH','ZXLM','ZXLR','ZXMS','ZXMU') THEN 'UPH'
        WHEN a1.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB','ZDBC','ZABC','ZECD') THEN 'CG'
        WHEN a1.ITCLS IN ('ZKIS') THEN 'Bedding'		
        WHEN a1.ITCLS IN ('WPLS') THEN 'Plastics'
        WHEN a1.ITCLS IN ('WVBC','WVCS') THEN 'Foundation'		
		WHEN a1.ITCLS IN ('PANL') THEN 'Panel'
		WHEN a1.ITCLS IN ('ZKIZ') THEN 'ZipperCover'
        WHEN a1.ITCLS IN ('BBFR','WVHC') THEN 'Verona'		
		WHEN a1.ITCLS NOT LIKE 'Z%' THEN 'Raw'
        ELSE 'Check' END) AS Product,
a1.ITMCQTY, b1.BXDCOMPONENTITEMNUMBER as BOX_MRP, b1.BXDCOMPONENTITEMDESCRIPTION as BOX_DES
FROM
(
-- MIL Inv Qty
SELECT t1.ITNBR, t2.ITDSC, t2.ITCLS, t1.HOUSE, t1.WHSLC, t1.MOHTQ,t2.B2Z95S,t3.ITMCQTY
FROM AMFLIBL.ITEMBL t1 left join AMFLIBL.ITMRVA t2 on t1.itnbr = t2.itnbr 
LEFT JOIN AFILELIBL.ITMEXT t3 on t1.itnbr = t3.itnbr
WHERE t1.WHSLC IN ('FA00') AND TRIM(t1.ITNBR) LIKE '%UN'
) a1
LEFT JOIN 
-- GET BOX_MRP and BOX_DES FROM BOM
(SELECT T1.BXDSTID,T1.BXDPARENTITEMNUMBER,T1.BXDCOMPONENTITEMNUMBER, T1.BXDCOMPONENTITEMDESCRIPTION
FROM RGNFILL.PSTBOMD T1
Where  T1.BXDSTID IN ('51') AND TRIM(T1.BXDPARENTITEMNUMBER) LIKE '%UN' AND  
(CASE WHEN T1.BXDCOMPONENTITEMDESCRIPTION LIKE 'RSC%' THEN 1 
	  WHEN T1.BXDUNITVOLUME >0 THEN 1 
	  WHEN T1.BXDCOMPONENTITEMDESCRIPTION LIKE 'UWALL%' THEN 1 
	  WHEN T1.BXDCOMPONENTITEMDESCRIPTION LIKE '%X%X%INCH%' THEN 1
	  ELSE 0 END)=1) AS b1
ON a1.HOUSE = b1.BXDSTID AND TRIM(a1.ITNBR)= TRIM(b1.BXDPARENTITEMNUMBER)
ORDER BY a1.ITNBR