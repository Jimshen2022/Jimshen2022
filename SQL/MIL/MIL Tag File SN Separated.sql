-- tag file serial number separated on Dec.16.2021
DECLARE SN_SHIPPED AS VARIANCE
SN_SHIPPED = SELECT CHAR(TRIM(A.WCSSERIALNUMBER)) AS SN FROM LLUSAF.WVCNTSD A, LLUSAF.WVCNTHD B WHERE A.WCSCONTAINERNUMBER = B.WCHCONTAINERNUMBER AND B.WCHCONTAINERSTATUS IN ('P','T') AND B.WCHPOSTEDTIMESTAMP BETWEEN CHAR(CURRENT DATE - 61 DAYS) AND CHAR(CURRENT DATE + 1 DAYS)



SELECT CHAR(TRIM(T1.TDTAG#)) AS SN,T1.TDITEM,T1.TDAPO# AS MO,T1.TDWHSE,T1.TDMDAT,T1.TDMTME, T1.TDTSTS,RIGHT('000000'||LTRIM(T1.TDMTME),6) AS TXT_TIME,T3.ITCLS,T5.ITMCQTY, 
(CASE 
	WHEN T3.ITCLS IN ('WPLS') THEN 'PLASTICS' 
	WHEN T3.ITCLS IN ('WVBC','WVHC') THEN 'FOUNDATION' 
	WHEN T3.ITCLS IN ('SLDK') THEN 'RP' 
	WHEN T3.ITCLS LIKE 'T%' THEN 'RP' 
	WHEN T3.ITCLS IN ('ZKIS') THEN 'BEDDING' 
	WHEN T3.ITCLS IN ('ZKIZ') THEN 'ZIPPERCOVER' 
	WHEN T3.ITCLS LIKE 'Z%' AND T3.ITCLS LIKE '%K' THEN 'UNKITS' 
	WHEN T3.ITCLS IN ('PACS') THEN 'UNKITS' 
	WHEN T3.ITCLS IN ('BBFR') THEN 'FR SOCK' 
	WHEN T3.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB') THEN 'CG' 
	WHEN T3.ITCLS LIKE 'Z%' THEN 'UPH' ELSE 'CHECK' END) AS PRODUCT, 
(CASE 
	WHEN T1.TDMTME BETWEEN '070000' AND '194459' THEN 'DS' ELSE 'NS' END) AS SHIFT,
(CASE 
	WHEN  CHAR(TRIM(T1.TDTAG#)) IN 
		(SELECT CHAR(TRIM(A.WCSSERIALNUMBER)) AS SN FROM LLUSAF.WVCNTSD A, LLUSAF.WVCNTHD B WHERE A.WCSCONTAINERNUMBER = B.WCHCONTAINERNUMBER AND B.WCHCONTAINERSTATUS IN ('P','T') AND B.WCHPOSTEDTIMESTAMP BETWEEN CHAR(CURRENT DATE - 61 DAYS) AND CHAR(CURRENT DATE + 1 DAYS)) 
	THEN 'SHIPPED'
	ELSE 'ONSTAGE' END) AS TYPE
	
FROM DISTLIBL.TAGINVD T1,(SELECT DISTINCT T2.ITNBR,T2.ITCLS FROM AMFLIBL.ITEMBL T2 WHERE T2.HOUSE = '51' GROUP BY T2.ITNBR,T2.ITCLS) AS T3, 
(SELECT DISTINCT T4.ITNBR,T4.ITMCQTY FROM AFILELIBL.ITMEXT T4 GROUP BY T4.ITNBR,T4.ITMCQTY) AS T5 
WHERE T1.TDITEM = T3.ITNBR AND T1.TDITEM=T5.ITNBR AND T3.ITNBR=T5.ITNBR AND T1.TDTSTS IN ('R','S','') 
AND T1.TDMDAT BETWEEN  INT('1'||SUBSTR(TRIM(CHAR(CURRENT DATE - 10 DAYS)),3,2)||SUBSTR(TRIM(CHAR(CURRENT DATE- 10 DAYS)),6,2)||SUBSTR(TRIM(CHAR(CURRENT DATE- 10 DAYS)),9,2))  
AND INT('1'||SUBSTR(TRIM(CHAR(CURRENT DATE + 1 DAYS)),3,2)||SUBSTR(TRIM(CHAR(CURRENT DATE + 1 DAYS)),6,2)||SUBSTR(TRIM(CHAR(CURRENT DATE + 1 DAYS)),9,2)) 
AND (T3.ITCLS LIKE 'Z%' AND T3.ITCLS NOT LIKE '%K' AND T3.ITCLS NOT IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB','ZKIS','ZKIZ')) 
