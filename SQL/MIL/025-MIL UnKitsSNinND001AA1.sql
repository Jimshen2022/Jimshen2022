-- 025-MIL UnKits SN in ND001AA1(no demand location), create by Jimshen on Jul.19.2022
SELECT *
FROM
(
-- By t2 to get AACOD1 
SELECT t1.AACOD1,t1.AAORD#,t1.AAITM#,t1.AATARA,CHAR(t1.AASER#) SN,t1.AAADAT,t1.AAATIM,t1.AAAUSR
FROM DISTLIBL.ACTAUDT t1
WHERE  t1.AAADAT BETWEEN (?-30) AND (?-1)
AND CHAR(rtrim(CHAR(t1.AASER#))||rtrim(t1.AAITM#)||rtrim(to_date(t1.AAADAT||' '||right('000000'||ltrim(t1.AAATIM),6), 'yyyymmdd hh24:mi:ss'))) IN 
	 -- Table t2 to Get 1365 max date time SN list 
	(SELECT CHAR(rtrim(t2.SN)||rtrim(t2.AAITM#)||rtrim(t2.ScannedTime)) AS SNSN
	FROM (
		 SELECT  CHAR(a.AASER#) SN,a.AAITM#, MAX(to_date(a.AAADAT||' '||right('000000'||ltrim(a.AAATIM),6), 'yyyymmdd hh24:mi:ss')) AS ScannedTime 
		 FROM  DISTLIBL.ACTAUDT a 
		 WHERE a.AASER#>0 AND a.AAADAT BETWEEN (?-30) AND (?-1)
		 GROUP BY CHAR(a.AASER#),a.AAITM#
		 )  t2
	 )
) x1	 
WHERE x1.AACOD1 IN ('MV') AND trim(x1.AAITM#) LIKE '%UN' AND trim(x1.AAITM#) NOT LIKE 'M%' AND x1.AATARA LIKE 'HJ%'
limit 10


-- 025-MIL UnKits SN in ND001AA1(no demand location), create by Jimshen on Jul.19.2022, BY QTY

SELECT x1.AACOD1,x1.AAORD#,x1.AAITM#,x1.AATARA,x1.AAADAT, x1.AAAUSR,a2.ITMCQTY, COUNT(x1.SN) as QTY
FROM
(
-- By Table t2 to get AACOD1.....by where conditions 
SELECT t1.AACOD1,t1.AAORD#,t1.AAITM#,t1.AATARA,CHAR(t1.AASER#) SN,t1.AAADAT,t1.AAATIM,t1.AAAUSR
FROM DISTLIBL.ACTAUDT t1
WHERE  t1.AAADAT BETWEEN (?-30) AND (?-1)
AND CHAR(rtrim(CHAR(t1.AASER#))||rtrim(t1.AAITM#)||rtrim(to_date(t1.AAADAT||' '||right('000000'||ltrim(t1.AAATIM),6), 'yyyymmdd hh24:mi:ss'))) IN 
	 -- Table t2 to Get 1365 max date time SN list 
	(SELECT CHAR(rtrim(t2.SN)||rtrim(t2.AAITM#)||rtrim(t2.ScannedTime)) AS SNSN
	FROM (
		 SELECT  CHAR(a.AASER#) SN,a.AAITM#, MAX(to_date(a.AAADAT||' '||right('000000'||ltrim(a.AAATIM),6), 'yyyymmdd hh24:mi:ss')) AS ScannedTime 
		 FROM  DISTLIBL.ACTAUDT a 
		 WHERE a.AASER#>0 AND a.AAADAT BETWEEN (?-30) AND (?-1)
		 GROUP BY CHAR(a.AASER#),a.AAITM#) t2)) x1	 
LEFT JOIN
-- to pull out data of Pices per Carton
(SELECT DISTINCT t4.ITNBR,t4.ITMCQTY FROM AFILELIBL.ITMEXT t4 GROUP BY t4.ITNBR,t4.ITMCQTY) AS a2 ON x1.AAITM# = a2.ITNBR
WHERE x1.AACOD1 IN ('MV') AND trim(x1.AAITM#) LIKE '%UN' AND trim(x1.AAITM#) NOT LIKE 'M%' AND x1.AATARA LIKE 'HJ%'
GROUP BY x1.AACOD1,x1.AAORD#,x1.AAITM#,x1.AATARA,x1.AAADAT, x1.AAAUSR,a2.ITMCQTY


