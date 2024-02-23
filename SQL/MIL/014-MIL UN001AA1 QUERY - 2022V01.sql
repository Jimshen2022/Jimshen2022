-- MIL UnKits last transaction on 1365 on Jun.01.2022
 
SELECT *

FROM
(
-- By Table t2 to get AACOD1 IN 'MV' to get which SN in 'UN%' location  
SELECT t1.AACOD1,t1.AAORD#,t1.AAITM#,t1.AATARA,CHAR(t1.AASER#) SN,t1.AAADAT,t1.AAATIM,t1.AAAUSR
FROM DISTLIBL.ACTAUDT t1
WHERE  t1.AAADAT BETWEEN CHAR(VARCHAR_FORMAT(current date -30 days,'YYYYMMDD'))  AND CHAR(VARCHAR_FORMAT(current date, 'YYYYMMDD'))
AND CHAR(rtrim(CHAR(t1.AASER#))||rtrim(t1.AAITM#)||rtrim(to_date(t1.AAADAT||' '||right('000000'||ltrim(t1.AAATIM),6), 'yyyymmdd hh24:mi:ss'))) IN 
	 -- Table t2 to Get 1365 max date time SN list --- last  
	(SELECT CHAR(rtrim(t2.SN)||rtrim(t2.AAITM#)||rtrim(t2.ScannedTime)) AS SNSN
	FROM (
		 SELECT  CHAR(a.AASER#) SN,a.AAITM#, MAX(to_date(a.AAADAT||' '||right('000000'||ltrim(a.AAATIM),6), 'yyyymmdd hh24:mi:ss')) AS ScannedTime 
		 FROM  DISTLIBL.ACTAUDT a 
		 WHERE a.AASER#>0 AND a.AAADAT BETWEEN CHAR(VARCHAR_FORMAT(current date -30 days,'YYYYMMDD'))  AND CHAR(VARCHAR_FORMAT(current date, 'YYYYMMDD'))
		 GROUP BY CHAR(a.AASER#),a.AAITM#
		 )  t2
	 )
) x1	 
WHERE x1.AACOD1 IN ('MV') AND trim(x1.AAITM#) LIKE '%UN' AND trim(x1.AAITM#) NOT LIKE 'M%' AND x1.AATARA LIKE 'UN%'
	 
	 
	 
	 