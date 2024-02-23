-- WN3 F/G EOL scanned time to loading time created by Jimshen on Aug.03.2022

--MO,EOL scanned time
SELECT a3.AACOD1,a3.AACOD2,a3.AAORD#,a3.AAITM#,a3.SN,a3.ITCLS,a3.Product,a3."EOL_Scanned_Time",a3."HJ_Received_Time", a4."HJ_Loading_Time",
(CASE WHEN a4."HJ_Loading_Time" IS NULL THEN 'NotLoading' ELSE 'Loaded' END) AS Status,
(CASE 
	WHEN a4."HJ_Loading_Time" IS NOT NULL THEN (values timestampdiff(16,char(a4."HJ_Loading_Time" - a3."EOL_Scanned_Time"))) 
	ELSE (values timestampdiff(16,char(CURRENT TIMESTAMP - a3."EOL_Scanned_Time"))) END)  AS "PendingDays",
(CASE 
	WHEN (CASE 
	WHEN a4."HJ_Loading_Time" IS NOT NULL THEN (values timestampdiff(16,char(a4."HJ_Loading_Time" - a3."EOL_Scanned_Time"))) 
	ELSE (values timestampdiff(16,char(CURRENT TIMESTAMP - a3."EOL_Scanned_Time"))) END) < 3 THEN 'a. 0-3 days'
	WHEN (CASE 
	WHEN a4."HJ_Loading_Time" IS NOT NULL THEN (values timestampdiff(16,char(a4."HJ_Loading_Time" - a3."EOL_Scanned_Time"))) 
	ELSE (values timestampdiff(16,char(CURRENT TIMESTAMP - a3."EOL_Scanned_Time"))) END) < 7 THEN 'b. 3-7 days'
	WHEN (CASE 
	WHEN a4."HJ_Loading_Time" IS NOT NULL THEN (values timestampdiff(16,char(a4."HJ_Loading_Time" - a3."EOL_Scanned_Time"))) 
	ELSE (values timestampdiff(16,char(CURRENT TIMESTAMP - a3."EOL_Scanned_Time"))) END) < 10 THEN 'c. 7-10 days'
    WHEN (CASE 
	WHEN a4."HJ_Loading_Time" IS NOT NULL THEN (values timestampdiff(16,char(a4."HJ_Loading_Time" - a3."EOL_Scanned_Time"))) 
	ELSE (values timestampdiff(16,char(CURRENT TIMESTAMP - a3."EOL_Scanned_Time"))) END) < 14 THEN 'd. 10-14 days'
	WHEN (CASE 
	WHEN a4."HJ_Loading_Time" IS NOT NULL THEN (values timestampdiff(16,char(a4."HJ_Loading_Time" - a3."EOL_Scanned_Time"))) 
	ELSE (values timestampdiff(16,char(CURRENT TIMESTAMP - a3."EOL_Scanned_Time"))) END) < 21 THEN 'e. 14-21 days'
	WHEN (CASE 
	WHEN a4."HJ_Loading_Time" IS NOT NULL THEN (values timestampdiff(16,char(a4."HJ_Loading_Time" - a3."EOL_Scanned_Time"))) 
	ELSE (values timestampdiff(16,char(CURRENT TIMESTAMP - a3."EOL_Scanned_Time"))) END) < 30 THEN 'f. 21-30 days'
	WHEN (CASE 
	WHEN a4."HJ_Loading_Time" IS NOT NULL THEN (values timestampdiff(16,char(a4."HJ_Loading_Time" - a3."EOL_Scanned_Time"))) 
	ELSE (values timestampdiff(16,char(CURRENT TIMESTAMP - a3."EOL_Scanned_Time"))) END) < 45 THEN 'g. 30-45 days'
	WHEN (CASE 
	WHEN a4."HJ_Loading_Time" IS NOT NULL THEN (values timestampdiff(16,char(a4."HJ_Loading_Time" - a3."EOL_Scanned_Time"))) 
	ELSE (values timestampdiff(16,char(CURRENT TIMESTAMP - a3."EOL_Scanned_Time"))) END) < 60 THEN 'h. 45-60 days'
	WHEN (CASE 
	WHEN a4."HJ_Loading_Time" IS NOT NULL THEN (values timestampdiff(16,char(a4."HJ_Loading_Time" - a3."EOL_Scanned_Time"))) 
	ELSE (values timestampdiff(16,char(CURRENT TIMESTAMP - a3."EOL_Scanned_Time"))) END) < 75 THEN 'i. 60-75 days'
    WHEN (CASE 
	WHEN a4."HJ_Loading_Time" IS NOT NULL THEN (values timestampdiff(16,char(a4."HJ_Loading_Time" - a3."EOL_Scanned_Time"))) 
	ELSE (values timestampdiff(16,char(CURRENT TIMESTAMP - a3."EOL_Scanned_Time"))) END) < 90 THEN 'j. 75-90 days'
	WHEN (CASE 
	WHEN a4."HJ_Loading_Time" IS NOT NULL THEN (values timestampdiff(16,char(a4."HJ_Loading_Time" - a3."EOL_Scanned_Time"))) 
	ELSE (values timestampdiff(16,char(CURRENT TIMESTAMP - a3."EOL_Scanned_Time"))) END) < 120 THEN 'k. 90-120 days'
	WHEN (CASE 
	WHEN a4."HJ_Loading_Time" IS NOT NULL THEN (values timestampdiff(16,char(a4."HJ_Loading_Time" - a3."EOL_Scanned_Time"))) 
	ELSE (values timestampdiff(16,char(CURRENT TIMESTAMP - a3."EOL_Scanned_Time"))) END) < 150 THEN 'l. 120-150 days'
	WHEN (CASE 
	WHEN a4."HJ_Loading_Time" IS NOT NULL THEN (values timestampdiff(16,char(a4."HJ_Loading_Time" - a3."EOL_Scanned_Time"))) 
	ELSE (values timestampdiff(16,char(CURRENT TIMESTAMP - a3."EOL_Scanned_Time"))) END) < 180 THEN 'm. 150-180 days'
	WHEN (CASE 
	WHEN a4."HJ_Loading_Time" IS NOT NULL THEN (values timestampdiff(16,char(a4."HJ_Loading_Time" - a3."EOL_Scanned_Time"))) 
	ELSE (values timestampdiff(16,char(CURRENT TIMESTAMP - a3."EOL_Scanned_Time"))) END) < 270 THEN 'n. 180-270 days'
	WHEN (CASE 
	WHEN a4."HJ_Loading_Time" IS NOT NULL THEN (values timestampdiff(16,char(a4."HJ_Loading_Time" - a3."EOL_Scanned_Time"))) 
	ELSE (values timestampdiff(16,char(CURRENT TIMESTAMP - a3."EOL_Scanned_Time"))) END) < 360 THEN 'o. 270-360 days'
	ELSE 'p. Over 360 days' END) AS "LoadingTime - EOLScannedTime"
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
FROM DISTLIBW.ACTAUDT t1, (SELECT TRIM(T2.ITNBR) ITNBR, T2.ITCLS FROM AMFLIBW.ITEMBL T2 WHERE T2.HOUSE IN ('35')) t3
WHERE  t1.AAADAT BETWEEN  CHAR(VARCHAR_FORMAT(current date - 90 days,'YYYYMMDD'))  AND CHAR(VARCHAR_FORMAT(current date, 'YYYYMMDD')) 
 AND t1.AACOD1 IN ('MF') AND t1.AACOD2 IN ('SN') AND TRIM(t1.AAITM#) = t3.ITNBR AND t1.AATWHS IN ('35')
 ORDER BY to_date(t1.AAADAT||' '||right('000000'||ltrim(t1.AAATIM),6), 'yyyymmdd hh24:mi:ss')
) a1

LEFT JOIN 
(
--HJ RECEIVED TIME
SELECT t1.AACOD1,t1.AACOD2,t1.AAORD#,TRIM(t1.AAITM#) AAITM#,CHAR(t1.AASER#) as SN,to_date(t1.AAADAT||' '||right('000000'||ltrim(t1.AAATIM),6), 'yyyymmdd hh24:mi:ss') as "HJ_Received_Time", t1.AAAPGM
FROM DISTLIBW.ACTAUDT t1
WHERE  t1.AAADAT BETWEEN  CHAR(VARCHAR_FORMAT(current date -30 days,'YYYYMMDD'))  AND CHAR(VARCHAR_FORMAT(current date, 'YYYYMMDD')) 
 AND t1.AACOD1 IN ('MF') AND t1.AACOD2 IN ('IT') AND t1.AATWHS IN ('35')
) a2 ON a1.SN = a2.SN
) a3

LEFT JOIN
(
--HJ Loading TIME
SELECT t1.AACOD1,t1.AACOD2,TRIM(t1.AAITM#) AAITM#,CHAR(t1.AASER#) as SN, max(to_date(t1.AAADAT||' '||right('000000'||ltrim(t1.AAATIM),6), 'yyyymmdd hh24:mi:ss')) as "HJ_Loading_Time" 
FROM DISTLIBW.ACTAUDT t1
WHERE  t1.AAADAT BETWEEN  CHAR(VARCHAR_FORMAT(current date - 90 days,'YYYYMMDD'))  AND CHAR(VARCHAR_FORMAT(current date, 'YYYYMMDD')) 
 AND t1.AACOD1 IN ('BT') AND t1.AAAPGM IN ('HJ361E') AND t1.AAFWHS IN ('35')
GROUP BY t1.AACOD1,t1.AACOD2,TRIM(t1.AAITM#),CHAR(t1.AASER#)
) a4 ON a3.SN=A4.SN