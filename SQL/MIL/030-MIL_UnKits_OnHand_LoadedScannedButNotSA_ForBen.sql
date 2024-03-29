-- MIL_UnKits_OnHand_LoadedScannedButNotSA_ForBen

-- MIL AS400 STOCK BY CARTONS  - 20211210 
SELECT x1.ITNBR, x1.ITDSC, x1.ITCLS, x1.Product,x1.WH51_FA00, x1.WH51_Non_FA00, x1.WH52_FA00, x1.WH52_Non_FA00,
SUM(CASE WHEN y1.WCIDESTINATION IN ('1') THEN y1.LOADEDCARTONS ELSE 0 END) AS #1,
SUM(CASE WHEN y1.WCIDESTINATION IN ('5') THEN y1.LOADEDCARTONS ELSE 0 END) AS #5,
SUM(CASE WHEN y1.WCIDESTINATION IN ('ECR') THEN y1.LOADEDCARTONS ELSE 0 END) AS #ECR,
SUM(CASE WHEN y1.WCIDESTINATION IN ('12') THEN y1.LOADEDCARTONS ELSE 0 END) AS #12,
SUM(CASE WHEN y1.WCIDESTINATION IN ('15') THEN y1.LOADEDCARTONS ELSE 0 END) AS #15,
SUM(CASE WHEN y1.WCIDESTINATION IN ('17') THEN y1.LOADEDCARTONS ELSE 0 END) AS #17,
SUM(CASE WHEN y1.WCIDESTINATION IN ('19') THEN y1.LOADEDCARTONS ELSE 0 END) AS #19,
SUM(CASE WHEN y1.WCIDESTINATION IN ('20') THEN y1.LOADEDCARTONS ELSE 0 END) AS #20,
SUM(CASE WHEN y1.WCIDESTINATION IN ('28') THEN y1.LOADEDCARTONS ELSE 0 END) AS #28,
SUM(CASE WHEN y1.WCIDESTINATION IN ('42') THEN y1.LOADEDCARTONS ELSE 0 END) AS #42,
SUM(CASE WHEN y1.WCIDESTINATION IN ('001') THEN y1.LOADEDCARTONS ELSE 0 END) AS #001,
SUM(CASE WHEN y1.WCIDESTINATION IN ('101') THEN y1.LOADEDCARTONS ELSE 0 END) AS #101,
SUM(CASE WHEN y1.WCIDESTINATION IN ('103') THEN y1.LOADEDCARTONS ELSE 0 END) AS #103
FROM
(
SELECT t2.ITNBR, t1.ITDSC, t1.ITCLS,
(CASE 
        WHEN t1.ITCLS LIKE 'TAF%' THEN 'RP'
        WHEN t1.ITCLS IN ('PACS') THEN 'UnKits'
        WHEN t1.ITCLS LIKE 'Z%' AND t1.ITCLS LIKE '%K' THEN 'UnKits'
        WHEN t1.ITCLS IN ('ZACM','ZASU','ZMLH','ZMLR','ZUSR','ZUSU','ZVUC','ZXUC','ZUSU','ZUMU','ZAMU','ZASM','ZASR','ZDMA','ZMUC','ZSUS','ZUMS','ZUSM','ZVMA','ZVUS','ZXLH','ZXLM','ZXLR','ZXMS','ZXMU') THEN 'UPH'
        WHEN t1.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB','ZDBC','ZABC','ZECD') THEN 'CG'
        WHEN t1.ITCLS IN ('ZKIS') THEN 'Bedding'		
        WHEN t1.ITCLS IN ('WPLS') THEN 'Plastics'
        WHEN t1.ITCLS IN ('WVBC','WVCS') THEN 'Foundation'		
        WHEN t1.ITCLS IN ('PANL') THEN 'Panel'
        WHEN t1.ITCLS IN ('ZKIZ') THEN 'ZipperCover'
        WHEN t1.ITCLS IN ('BBFR','WVHC') THEN 'Verona'		
        WHEN t1.ITCLS NOT LIKE 'Z%' THEN 'RawMaterial'
        ELSE 'Check' END) AS Product,
       SUM(CASE WHEN t2.LLOCN IN ('FA00') AND t2.HOUSE in ('51') THEN CEIL(t2.LQNTY/t3.ITMCQTY) ELSE 0 END) AS WH51_FA00,
       SUM(CASE WHEN t2.LLOCN NOT IN ('FA00') AND t2.HOUSE in ('51') THEN CEIL(t2.LQNTY/t3.ITMCQTY) ELSE 0 END) AS WH51_Non_FA00,
	   SUM(CASE WHEN t2.LLOCN IN ('FA00') AND t2.HOUSE in ('52') THEN CEIL(t2.LQNTY/t3.ITMCQTY) ELSE 0 END) AS WH52_FA00,
       SUM(CASE WHEN t2.LLOCN NOT IN ('FA00') AND t2.HOUSE in ('52') THEN CEIL(t2.LQNTY/t3.ITMCQTY) ELSE 0 END) AS WH52_Non_FA00
	   
FROM AMFLIBL.ITEMASA t1, AMFLIBL.SLQNTY t2, AFILELIBL.ITMEXT t3
WHERE t1.ITNBR = t2.ITNBR and t2.ITNBR = t3.ITNBR 
GROUP BY t2.ITNBR, t1.ITDSC, t1.ITCLS
HAVING (t1.ITCLS LIKE 'Z%' AND t1.ITCLS LIKE '%K') or t1.ITCLS IN ('PACS')
ORDER BY t2.ITNBR
) x1

FULL OUTER JOIN 
-- Loading scanned but still not SA container Pieces
(
SELECT a.WCIDESTINATION, a.WCICONTAINERNUMBER AS Container#, a.WCIITEMNUMBER, SUM(CEIL(a.WCIQUANTITYLOADED/c.ITMCQTY)) AS LOADEDCARTONS,
	(CASE 
		WHEN a.WCICONTAINERNUMBER LIKE 'MRUN%' THEN 'InTempCTN'
		WHEN a.WCICONTAINERNUMBER LIKE 'KECR%' THEN 'InTempCTN'
		WHEN a.WCICONTAINERNUMBER LIKE 'KHO%' THEN 'InTempCTN'
		WHEN a.WCICONTAINERNUMBER LIKE 'M3K%' THEN 'InTempCTN'
		WHEN a.WCICONTAINERNUMBER LIKE 'M3E%' THEN 'InTempCTN'
		WHEN a.WCICONTAINERNUMBER LIKE 'M3H%' THEN 'InTempCTN'
		WHEN a.WCICONTAINERNUMBER LIKE 'RUN%' THEN 'InTempCTN'
		ELSE 'InRealCTN' END) AS "CTN_Status"
FROM DISTLIBL.TBL_WVCONTAINER_DTL_ITM a, DISTLIBL.TBL_WVCONTAINER_HDR b, AFILELIBL.ITMEXT c
WHERE a.WCICONTAINERNUMBER = b.WCHCONTAINERNUMBER AND a.WCIITEMNUMBER = c.ITNBR
AND b.WCHCONTAINERSTATUS NOT IN ('P','T') 
GROUP BY a.WCIDESTINATION, a.WCICONTAINERNUMBER, a.WCIITEMNUMBER,	(CASE 
		WHEN a.WCICONTAINERNUMBER LIKE 'MRUN%' THEN 'InTempCTN'
		WHEN a.WCICONTAINERNUMBER LIKE 'KECR%' THEN 'InTempCTN'
		WHEN a.WCICONTAINERNUMBER LIKE 'KHO%' THEN 'InTempCTN'
		WHEN a.WCICONTAINERNUMBER LIKE 'M3K%' THEN 'InTempCTN'
		WHEN a.WCICONTAINERNUMBER LIKE 'M3E%' THEN 'InTempCTN'
		WHEN a.WCICONTAINERNUMBER LIKE 'M3H%' THEN 'InTempCTN'
		WHEN a.WCICONTAINERNUMBER LIKE 'RUN%' THEN 'InTempCTN'
		ELSE 'InRealCTN' END)
HAVING (CASE 
		WHEN a.WCICONTAINERNUMBER LIKE 'MRUN%' THEN 'InTempCTN'
		WHEN a.WCICONTAINERNUMBER LIKE 'KECR%' THEN 'InTempCTN'
		WHEN a.WCICONTAINERNUMBER LIKE 'KHO%' THEN 'InTempCTN'
		WHEN a.WCICONTAINERNUMBER LIKE 'M3K%' THEN 'InTempCTN'
		WHEN a.WCICONTAINERNUMBER LIKE 'M3E%' THEN 'InTempCTN'
		WHEN a.WCICONTAINERNUMBER LIKE 'M3H%' THEN 'InTempCTN'
		WHEN a.WCICONTAINERNUMBER LIKE 'RUN%' THEN 'InTempCTN'
		ELSE 'InRealCTN' END) IN ('InRealCTN')
) y1  ON x1.ITNBR = y1.WCIITEMNUMBER
GROUP BY  x1.ITNBR, x1.ITDSC, x1.ITCLS, x1.Product,x1.WH51_FA00, x1.WH51_Non_FA00, x1.WH52_FA00, x1.WH52_Non_FA00
ORDER BY x1.ITNBR