-- MIL UnKits Received SN By Shift create by Jimshen on Jan.03.2022

SELECT  CHAR(a.AASER#) as SN,a.AACOD1,a.AATWHS,a.AATARA||'00'||a.AATASL||a.AATSEC||a.AATTIR as Location, a.AAORD#,a.AAITM#,
a.AAEMP#,a.AAAUSR,
(CASE WHEN char(substr(char(TIME(to_date(a.AAADAT||' '||right('000000'||ltrim(a.AAATIM),6), 'yyyymmdd hh24:mi:ss'))),1,2)||substr(char(TIME(to_date(a.AAADAT||' '||right('000000'||ltrim(a.AAATIM),6), 'yyyymmdd hh24:mi:ss'))),4,2)||
substr(char(TIME(to_date(a.AAADAT||' '||right('000000'||ltrim(a.AAATIM),6), 'yyyymmdd hh24:mi:ss'))),7,2)) BETWEEN '070000' AND '194459' THEN 'DS' ELSE 'NS' END) AS SHIFT,t2.ITMCQTY,
MIN(to_date(a.AAADAT||' '||right('000000'||ltrim(a.AAATIM),6), 'yyyymmdd hh24:mi:ss')) as ScannedTime,1/t2.ITMCQTY as Cartons

FROM  DISTLIBL.ACTAUDT a, (SELECT DISTINCT t4.ITNBR,t4.ITMCQTY FROM AFILELIBL.ITMEXT t4 GROUP BY t4.ITNBR,t4.ITMCQTY) AS t2
WHERE a.AACOD1 IN ('MV') and a.AASER#>0 and trim(a.AAITM#) LIKE '%UN' AND trim(a.AAITM#) NOT LIKE 'M%' AND a.AATARA LIKE 'HJ%'
and a.AAADAT BETWEEN ? AND ? AND a.AAITM# = t2.ITNBR AND NOT EXISTS 
(SELECT 1 FROM  DISTLIBL.ACTAUDT b WHERE b.AACOD1 IN ('MV') and b.AASER#>0 and (trim(b.AAITM#) LIKE '%UN' AND trim(b.AAITM#) NOT LIKE 'M%') AND b.AATARA LIKE 'HJ%'
and b.AAADAT BETWEEN (?-30) AND (?-1) and a.AASER#=b.AASER#)

GROUP BY CHAR(a.AASER#), a.AACOD1,a.AATWHS,a.AATARA||'00'||a.AATASL||a.AATSEC||a.AATTIR,a.AAORD#,a.AAITM#, a.AAEMP#,a.AAAUSR, (CASE WHEN char(substr(char(TIME(to_date(a.AAADAT||' '||right('000000'||ltrim(a.AAATIM),6), 'yyyymmdd hh24:mi:ss'))),1,2)||substr(char(TIME(to_date(a.AAADAT||' '||right('000000'||ltrim(a.AAATIM),6), 'yyyymmdd hh24:mi:ss'))),4,2)||
substr(char(TIME(to_date(a.AAADAT||' '||right('000000'||ltrim(a.AAATIM),6), 'yyyymmdd hh24:mi:ss'))),7,2)) BETWEEN '070000' AND '194459' THEN 'DS' ELSE 'NS' END),t2.ITMCQTY


-- MIL UnKits Received Carton By Shift create by Jimshen on Jan.03.2022


SELECT t1.AACOD1,t1.AATWHS,t1.AACOD1,t1.AATWHS, t1.Location, t1.AAORD#, t1.AAITM#, t1.AAEMP#, t1.AAAUSR, t1.SHIFT, DATE(t1.ScannedTime) AS SCANNEDDATE,t2.ITMCQTY, CEIL(COUNT(t1.SN)/t2.ITMCQTY) AS Cartons 
FROM (
SELECT  CHAR(a.AASER#) as SN,a.AACOD1,a.AATWHS,a.AATARA||'00'||a.AATASL||a.AATSEC||a.AATTIR as Location, a.AAORD#,a.AAITM#,
a.AAEMP#,a.AAAUSR,
(CASE WHEN char(substr(char(TIME(to_date(a.AAADAT||' '||right('000000'||ltrim(a.AAATIM),6), 'yyyymmdd hh24:mi:ss'))),1,2)||substr(char(TIME(to_date(a.AAADAT||' '||right('000000'||ltrim(a.AAATIM),6), 'yyyymmdd hh24:mi:ss'))),4,2)||
substr(char(TIME(to_date(a.AAADAT||' '||right('000000'||ltrim(a.AAATIM),6), 'yyyymmdd hh24:mi:ss'))),7,2)) BETWEEN '070000' AND '194459' THEN 'DS' ELSE 'NS' END) AS SHIFT,
MIN(to_date(a.AAADAT||' '||right('000000'||ltrim(a.AAATIM),6), 'yyyymmdd hh24:mi:ss')) as ScannedTime

FROM  DISTLIBL.ACTAUDT a
WHERE a.AACOD1 IN ('MV') and a.AASER#>0 and trim(a.AAITM#) LIKE '%UN' AND trim(a.AAITM#) NOT LIKE 'M%' AND a.AATARA LIKE 'HJ%'
and a.AAADAT BETWEEN ? AND ? AND NOT EXISTS 
(SELECT 1 FROM  DISTLIBL.ACTAUDT b WHERE b.AACOD1 IN ('MV') and b.AASER#>0 and (trim(b.AAITM#) LIKE '%UN' AND trim(b.AAITM#) NOT LIKE 'M%') AND b.AATARA LIKE 'HJ%'
and b.AAADAT BETWEEN (?-30) AND (?-1) and a.AASER#=b.AASER#)

GROUP BY CHAR(a.AASER#), a.AACOD1,a.AATWHS,a.AATARA||'00'||a.AATASL||a.AATSEC||a.AATTIR,a.AAORD#,a.AAITM#, a.AAEMP#,a.AAAUSR, (CASE WHEN char(substr(char(TIME(to_date(a.AAADAT||' '||right('000000'||ltrim(a.AAATIM),6), 'yyyymmdd hh24:mi:ss'))),1,2)||substr(char(TIME(to_date(a.AAADAT||' '||right('000000'||ltrim(a.AAATIM),6), 'yyyymmdd hh24:mi:ss'))),4,2)||
substr(char(TIME(to_date(a.AAADAT||' '||right('000000'||ltrim(a.AAATIM),6), 'yyyymmdd hh24:mi:ss'))),7,2)) BETWEEN '070000' AND '194459' THEN 'DS' ELSE 'NS' END)
) t1, (SELECT DISTINCT t4.ITNBR,t4.ITMCQTY FROM AFILELIBL.ITMEXT t4 GROUP BY t4.ITNBR,t4.ITMCQTY) AS t2
WHERE t1.AAITM# = t2.ITNBR
GROUP BY t1.AACOD1,t1.AATWHS,t1.AACOD1,t1.AATWHS, t1.Location, t1.AAORD#, t1.AAITM#, t1.AAEMP#, t1.AAAUSR, t1.SHIFT, DATE(t1.ScannedTime),t2.ITMCQTY



























