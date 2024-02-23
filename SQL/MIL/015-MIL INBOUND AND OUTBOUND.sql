
-- MIL Inbound Trx of Item class not like 'Z%'
SELECT t1.HOUSE,t1.TCODE,t1.ORDNO,t1.ITNBR,t2.ITCLS, t1.UPDDT,t1.UPDTM,t1.TRQTY,t1.ENTUM,t1.VNDNR,t1.REFNO,t1.LLOCN,t1.BATCH,t1.TRMID,
CHAR(t1.UPDDT||' '||right('000000'||ltrim(t1.UPDTM),6)) AS TrxTime, CHAR(SUBSTR(right('000000'||ltrim(t1.UPDTM),6),1,2)) AS HOUR,
(CASE 
        WHEN t2.ITCLS LIKE 'TAF%' THEN 'RP'
		WHEN t2.ITCLS IN ('PACS') THEN 'UnKits'
        WHEN t2.ITCLS LIKE 'Z%' AND t2.ITCLS LIKE '%K' THEN 'UnKits'
        WHEN t2.ITCLS IN ('ZACM','ZASU','ZMLH','ZMLR','ZUSR','ZUSU','ZVUC','ZXUC','ZUSU','ZUMU') THEN 'UPH'
        WHEN t2.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB') THEN 'CG'
        WHEN t2.ITCLS IN ('ZKIS') THEN 'Bedding'		
        WHEN t2.ITCLS IN ('WPLS') THEN 'Plastics'
        WHEN t2.ITCLS IN ('WVBC','WVCS') THEN 'Foundation'		
		WHEN t2.ITCLS IN ('PANL') THEN 'Panel'
		WHEN t2.ITCLS IN ('ZKIZ') THEN 'ZipperCover'
        WHEN t2.ITCLS IN ('BBFR','WVHC') THEN 'Verona'		
		WHEN t2.ITCLS NOT LIKE 'Z%' THEN 'RawMaterial'
        ELSE 'Check' END) AS Product
		
FROM AMFLIBL.IMHIST  t1, AMFLIBL.ITMRVA t2, AMFLIBL.WHSMST t3
WHERE t1.ITNBR=t2.ITNBR  AND t2.STID = t3.STID AND t1.HOUSE = t3.WHID AND t1.TRQTY > 0 AND t1.TCODE IN ('RP','RM','PQ') AND 
CHAR(t1.UPDDT||' '||right('000000'||ltrim(t1.UPDTM),6)) BETWEEN CHAR('1'||VARCHAR_FORMAT(current date - 1 days,'yymmdd hh24:mi:ss'))  AND CHAR('1'||VARCHAR_FORMAT(current timestamp, 'yymmdd hh24:mi:ss'))
AND t2.ITCLS NOT LIKE 'Z%'

