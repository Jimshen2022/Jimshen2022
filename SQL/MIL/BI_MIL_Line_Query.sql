-- pull out the MO and Line information
SELECT Distinct(CHAR(TRIM(substr(x1.JOBNO,1,5)))) AS LINE
From
(
(SELECT t1.ORDNO,t1.JOBNO,t1.FITEM FROM AMFLIBL.MOMAST t1 where t1.CRDT
BETWEEN  int('1'||substr(trim(char(CURRENT DATE - 180 days)),3,2)||substr(trim(char(CURRENT DATE- 180 days)),6,2)||substr(trim(char(CURRENT DATE- 180 days)),9,2)) 
AND int('1'||substr(trim(char(CURRENT DATE + 1 days)),3,2)||substr(trim(char(CURRENT DATE + 1 days)),6,2)||substr(trim(char(CURRENT DATE + 1 days)),9,2))
order by t1.CRDT DESC) 

UNION ALL

(SELECT t1.ORDNO,t1.JOBNO,t1.FITEM FROM AMFLIBL.MOHMST t1 where t1.CRDT
BETWEEN  int('1'||substr(trim(char(CURRENT DATE - 180 days)),3,2)||substr(trim(char(CURRENT DATE- 180 days)),6,2)||substr(trim(char(CURRENT DATE- 180 days)),9,2)) 
AND int('1'||substr(trim(char(CURRENT DATE + 1 days)),3,2)||substr(trim(char(CURRENT DATE + 1 days)),6,2)||substr(trim(char(CURRENT DATE + 1 days)),9,2))
order by t1.CRDT DESC)
) as x1
Group by CHAR(TRIM(SUBSTR(x1.JOBNO,1,5)))
order by LINE
