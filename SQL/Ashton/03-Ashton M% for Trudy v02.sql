
-- MIL EOL SCANNED M% SN
SELECT T1.AAAPGM,T1.AACOD1,T1.AACOD2,T1.AAORD#,T1.AAITM#,CHAR(T1.AASER#) AS SN,T1.AATQTY,T1.AAADAT,T1.AAATIM,T1.AATWHS
FROM DISTLIBL.ACTAUDT T1
WHERE T1.AATWHS IN ('51') AND T1.AATARA = 'UC' AND T1.AAITM# LIKE 'M%' AND 
T1.AAADAT BETWEEN '20180101' AND '20291231'


-- ASHTON PO RECEIVED SN
SELECT W.SN,W.AAITM#,W.AAEQP#,W.AAADAT,W.AAATIM,W.AATWHS,W.AATFR#
FROM (
-- table W include current SN received and archived SN received of M%
(SELECT DISTINCT CHAR(T2.AASER#) AS SN,T2.AAITM#,T2.AAEQP#,T2.AAADAT,T2.AAATIM,T2.AATWHS,T2.AATFR#
FROM DISTLIB.ACTAUDT T2
WHERE T2.AAAPGM IN ('HJPOTOWA','HJ157E') AND T2.AACOD1 IN ('RC') AND T2.AACOD2 IN ('SN') AND T2.AATWHS IN ('335')
AND T2.AATARA IN ('AC') AND T2.AAITM# LIKE 'M%' AND T2.AAADAT BETWEEN '20190601' AND '20291231')
UNION ALL
(SELECT DISTINCT CHAR(T2.AASER#) AS SN,T2.AAITM#,T2.AAEQP#,T2.AAADAT,T2.AAATIM,T2.AATWHS,T2.AATFR#
FROM DISTLIBH.ACTAUDT T2
WHERE T2.AAAPGM IN ('HJPOTOWA','HJ157E') AND T2.AACOD1 IN ('RC') AND T2.AACOD2 IN ('SN') AND T2.AATWHS IN ('335')
AND T2.AATARA IN ('AC') AND T2.AAITM# LIKE 'M%' AND T2.AAADAT BETWEEN '20190601' AND '20291231')
) AS W
WHERE NOT EXISTS
	-- Remove BT transactons that be picked SN in current files
	 (SELECT 1
	  FROM (
	      SELECT T8.SN
		  FROM 
		  (SELECT CHAR(a.AASER#) AS SN, a.AAAPGM,a.AACOD1,a.AACOD2,a.AACOD3
		  FROM DISTLIB.ACTAUDT a
		  WHERE a.AAAPGM IN ('HJ321E') AND a.AACOD1 IN ('BT') AND a.AACOD2 IN ('SN') AND a.AACOD3 IN ('HJ','UL') AND a.AAFWHS IN ('335','')  AND a.AAITM# LIKE 'M%' AND a.AAADAT BETWEEN '20190101' AND '20291231' AND  
		  to_date(a.AAADAT||' '||right('000000'||ltrim(a.AAATIM),6), 'yyyymmdd hh24:mi:ss') IN 
		  -- To find out the max datetime on BT Transactions 
		  (SELECT MAX(to_date(aa.AAADAT||' '||right('000000'||ltrim(aa.AAATIM),6), 'yyyymmdd hh24:mi:ss')) AS MAXDATE
		  FROM DISTLIB.ACTAUDT aa
		  WHERE aa.AAAPGM IN ('HJ321E') AND aa.AACOD1 IN ('BT') AND aa.AACOD2 IN ('SN') AND aa.AACOD3 IN ('HJ','UL') AND aa.AAFWHS IN ('335','')  AND aa.AAITM# LIKE 'M%' AND aa.AAADAT BETWEEN '20190101' AND '20291231'
		  GROUP BY CHAR(aa.AASER#))
		  ) T8
		  -- below where is to find BT transaction AACOD3 not equal to UL that means not belong to UNLOADING
		  WHERE T8.AACOD3 IN ('HJ')
		  ) T9
	 WHERE W.SN=T9.SN)
AND NOT EXISTS
	-- Remove BT transactons that be picked SN in archived files	
	 (SELECT 1
	  FROM (
	      SELECT T10.SN
		  FROM 
		  (SELECT CHAR(a.AASER#) AS SN, a.AAAPGM,a.AACOD1,a.AACOD2,a.AACOD3
		  FROM DISTLIBH.ACTAUDT a
		  WHERE a.AAAPGM IN ('HJ321E') AND a.AACOD1 IN ('BT') AND a.AACOD2 IN ('SN') AND a.AACOD3 IN ('HJ','UL') AND a.AAFWHS IN ('335','')  AND a.AAITM# LIKE 'M%' AND a.AAADAT BETWEEN '20190101' AND '20291231' AND  
		  to_date(a.AAADAT||' '||right('000000'||ltrim(a.AAATIM),6), 'yyyymmdd hh24:mi:ss') IN 
		  -- To find out the max datetime on BT Transactions 
		  (SELECT MAX(to_date(aa.AAADAT||' '||right('000000'||ltrim(aa.AAATIM),6), 'yyyymmdd hh24:mi:ss')) AS MAXDATE
		  FROM DISTLIBH.ACTAUDT aa
		  WHERE aa.AAAPGM IN ('HJ321E') AND aa.AACOD1 IN ('BT') AND aa.AACOD2 IN ('SN') AND aa.AACOD3 IN ('HJ','UL') AND aa.AAFWHS IN ('335','')  AND aa.AAITM# LIKE 'M%' AND aa.AAADAT BETWEEN '20190101' AND '20291231'
		  GROUP BY CHAR(aa.AASER#))
		  ) T10
		  -- below where is to find BT transaction AACOD3 not equal to UL that means not belong to UNLOADING
		  WHERE T10.AACOD3 IN ('HJ')
		  ) T12
	 WHERE W.SN=T12.SN)		  