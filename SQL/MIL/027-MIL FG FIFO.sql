-- 025-MIL UnKits SN in ND001AA1(no demand location), create by Jimshen on Jul.19.2022, BY QTY

SELECT x1.AACOD1,x1.AAORD#,x1.AAITM#,x1.AATARA,x1.SN,to_date(x1.AAADAT||' '||right('000000'||ltrim(x1.AAATIM),6), 'yyyymmdd hh24:mi:ss') as ScannedTime, x1.AAAUSR
FROM
(
-- By Table t2 to get AACOD1.....by where conditions 
SELECT t1.AACOD1,t1.AAORD#,t1.AAITM#,t1.AATARA,CHAR(t1.AASER#) SN,t1.AAADAT,t1.AAATIM,t1.AAAUSR
FROM DISTLIBL.ACTAUDT t1
WHERE  t1.AAADAT BETWEEN  '20210101' AND '20291231'
AND CHAR(rtrim(CHAR(t1.AASER#))||rtrim(t1.AAITM#)||rtrim(to_date(t1.AAADAT||' '||right('000000'||ltrim(t1.AAATIM),6), 'yyyymmdd hh24:mi:ss'))) IN 
	 -- Table t2 to Get 1365 min date time SN list 
	(SELECT CHAR(rtrim(t2.SN)||rtrim(t2.AAITM#)||rtrim(t2.ScannedTime)) AS SNSN
	FROM (
		 SELECT  CHAR(a.AASER#) SN,a.AAITM#, MIN(to_date(a.AAADAT||' '||right('000000'||ltrim(a.AAATIM),6), 'yyyymmdd hh24:mi:ss')) AS ScannedTime 
		 FROM  DISTLIBL.ACTAUDT a 
		 WHERE a.AASER#>0 AND a.AAADAT BETWEEN '20210101' AND '20291231'
		 GROUP BY CHAR(a.AASER#),a.AAITM#) t2)) x1	 
WHERE x1.AACOD1 IN ('MF') AND x1.AATARA LIKE 'UC%'





--SN RECEIVED DATE

SELECT b.AAITM#,b.SN, b.ScannedTime
FROM 
(
SELECT  a.AAITM#,CHAR(a.AASER#) SN, MIN(to_date(a.AAADAT||' '||right('000000'||ltrim(a.AAATIM),6), 'yyyymmdd hh24:mi:ss')) AS ScannedTime 
FROM  DISTLIBL.ACTAUDT a
WHERE a.AASER#>0 AND a.AACOD1 IN ('MF') AND a.AATARA LIKE 'UC%'
GROUP BY a.AAITM#,CHAR(a.AASER#)
) b 
WHERE EXISTS
-- MIL ON HAND ITEMS
(SELECT 1
 FROM 
 (SELECT t1.HOUSE, t1.ITNBR, t1.ITCLS, t1.MOHTQ FROM  AMFLIBL.ITEMBL t1  WHERE T1.ITCLS LIKE 'Z%' AND t1.ITCLS NOT LIKE '%K' AND t1.MOHTQ>0 
 ORDER BY t1.ITNBR) as x  WHERE b.AAITM# = x.ITNBR)   
		 
		 
-- OPEN CO		 
SELECT t1.CDA3CD as Wanek, t1.CDCVNB as Order_Number,t1.CDAITX as item, t1.CDD0NB,
Date('20'||Substr(t1.CDD0NB, 2, 2) || '-'||  Substr(t1.CDD0NB, 4, 2)|| '-' ||substr(t1.CDD0NB, 6, 2)) AS ETD,
t1.CDB9CD as Warehouse,t1.CDAGNV as Qty,t1.CDGLCD,t1.CDALTX,t1.CDAMDT,t1.CDAFVN,t1.CDAGVN
FROM AMFLIBL.MBCDRESM t1
WHERE t1.CDAGNV >0 and t1.CDGLCD like 'Z%' AND t1.CDGLCD not like '%K'



-- Francisco FIFO

--MO,EOL scanned time
SELECT a3.AACOD1,a3.AACOD2,a3.AAORD#,a3.AAITM#,a3.SN,a3.ITCLS,a3.Product,a3."EOL_Scanned_Time",a3."HJ_Received_Time", a4."HJ_Loading_Time" 
FROM 
(
SELECT a1.AACOD1,a1.AACOD2,a1.AAORD#,a1.AAITM#,a1.SN,a1.AAAPGM, a1.ITCLS,a1.Product,a1."EOL_Scanned_Time",a2."HJ_Received_Time"
FROM 
(SELECT t1.AACOD1,t1.AACOD2,t1.AAORD#,TRIM(t1.AAITM#) AAITM#,CHAR(t1.AASER#) as SN,to_date(t1.AAADAT||' '||right('000000'||ltrim(t1.AAATIM),6), 'yyyymmdd hh24:mi:ss') as "EOL_Scanned_Time",t1.AAAPGM, t3.ITCLS,
(CASE 
        WHEN t3.ITCLS LIKE 'TAF%' THEN 'RP'
		WHEN t3.ITCLS IN ('PACS') THEN 'UnKits'
        WHEN t3.ITCLS LIKE 'Z%' AND t3.ITCLS LIKE '%K' THEN 'UnKits'
        WHEN t3.ITCLS IN ('ZACM','ZASU','ZMLH','ZMLR','ZUSR','ZUSU','ZVUC','ZXUC','ZUSU','ZUMU') THEN 'UPH'
        WHEN t3.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB','ZDBC','ZABC','ZECD') THEN 'CG'
        WHEN t3.ITCLS IN ('ZKIS') THEN 'Bedding'		
        WHEN t3.ITCLS IN ('WPLS') THEN 'Plastics'
        WHEN t3.ITCLS IN ('WVBC','WVCS') THEN 'Foundation'		
		WHEN t3.ITCLS IN ('PANL') THEN 'Panel'
		WHEN t3.ITCLS IN ('ZKIZ') THEN 'ZipperCover'
        WHEN t3.ITCLS IN ('BBFR','WVHC') THEN 'Verona'		
		WHEN t3.ITCLS NOT LIKE 'Z%' THEN 'Raw'
        ELSE 'Check' END) AS Product
FROM DISTLIBL.ACTAUDT t1, (SELECT TRIM(T2.ITNBR) ITNBR, T2.ITCLS FROM AMFLIBL.ITEMBL T2 WHERE T2.HOUSE IN ('51')) t3
WHERE  t1.AAADAT BETWEEN  '20220801' AND '20291231' AND t1.AACOD1 IN ('MF') AND t1.AACOD2 IN ('SN') AND TRIM(t1.AAITM#) = t3.ITNBR
) a1

LEFT JOIN 
(
--HJ RECEIVED TIME
SELECT t1.AACOD1,t1.AACOD2,t1.AAORD#,TRIM(t1.AAITM#) AAITM#,CHAR(t1.AASER#) as SN,to_date(t1.AAADAT||' '||right('000000'||ltrim(t1.AAATIM),6), 'yyyymmdd hh24:mi:ss') as "HJ_Received_Time", t1.AAAPGM
FROM DISTLIBL.ACTAUDT t1
WHERE  t1.AAADAT BETWEEN  '20220801' AND '20291231' AND t1.AACOD1 IN ('MF') AND t1.AACOD2 IN ('IT')
) a2 ON a1.SN = a2.SN
) a3

LEFT JOIN
(
--HJ Loading TIME
SELECT t1.AACOD1,t1.AACOD2,TRIM(t1.AAITM#) AAITM#,CHAR(t1.AASER#) as SN, max(to_date(t1.AAADAT||' '||right('000000'||ltrim(t1.AAATIM),6), 'yyyymmdd hh24:mi:ss')) as "HJ_Loading_Time" 
FROM DISTLIBL.ACTAUDT t1
WHERE  t1.AAADAT BETWEEN  '20220801' AND '20291231' AND t1.AACOD1 IN ('BT') AND t1.AAAPGM IN ('HJ361E')
GROUP BY t1.AACOD1,t1.AACOD2,TRIM(t1.AAITM#),CHAR(t1.AASER#)
) a4 ON a3.SN=A4.SN 















































