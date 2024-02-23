
-- a1.MFIDT LoadDate  20201102
 to_date(char(a1.MFIDT),'yyyymmdd')  LoadDate,



--DB2日期与时间



DATE('20'||SUBSTR(CHAR(T1.UPDDT),2,2)||'-'||SUBSTR(CHAR(T1.UPDDT),4,2)||'-'||SUBSTR(CHAR(T1.UPDDT),6,2)) AS DATE,              
WEEK(DATE('20'||SUBSTR(CHAR(T1.UPDDT),2,2)||'-'||SUBSTR(CHAR(T1.UPDDT),4,2)||'-'||SUBSTR(CHAR(T1.UPDDT),6,2))) AS WEEK,
YEAR(DATE('20'||SUBSTR(CHAR(T1.UPDDT),2,2)||'-'||SUBSTR(CHAR(T1.UPDDT),4,2)||'-'||SUBSTR(CHAR(T1.UPDDT),6,2))) AS YEAR


--AND T1.UPDDT BETWEEN '1210101' AND '1211231'
AND T1.UPDDT BETWEEN CHAR('1'||VARCHAR_FORMAT(current date -30 days,'YYMMDD')) AND CHAR('1'||VARCHAR_FORMAT(current date,'YYMMDD'))



values VARCHAR_FORMAT(current TIMESTAMP,'yyyy-mm-dd hh24:mi:ss');   显示 2121-08-07

--两个日期时间间隔天数
CASE 
	WHEN a4."HJ_Loading_Time" IS NOT NULL THEN (values timestampdiff(16,char(a4."HJ_Loading_Time" - a3."EOL_Scanned_Time"))) 
	ELSE (values timestampdiff(16,char(CURRENT TIMESTAMP - a3."EOL_Scanned_Time"))) END)  AS "PendingDays",


-- GOOD TO SEE

RM Date From	5/30/2022 7:00:00	TO	6/1/2022 6:59:59

SELECT * FROM  DISTLIBL.ACTAUDT b WHERE b.AACOD1 IN ('MV') and b.AASER#>0 and (trim(b.AAITM#) LIKE '%UN' AND trim(b.AAITM#) NOT LIKE 'M%') AND b.AATARA LIKE 'HJ%'
	and to_date(b.AAADAT||' '||right('000000'||ltrim(b.AAATIM),6), 'yyyymmdd hh24:mi:ss') BETWEEN ? AND ?
limit 5



-- a.AAADAT LIKE 20220101 
SELECT  a.AAITM#, CHAR(a.AASER#) AS SN, to_date(a.AAADAT||' '||right('000000'||ltrim(a.AAATIM),6), 'yyyymmdd hh24:mi:ss') as RM_Time
FROM  DISTLIBL.ACTAUDT a 
 WHERE a.AAAPGM in ('HJ111E') AND  a.AAADAT Between CHAR(VARCHAR_FORMAT(current date -60 days,'YYYYMMDD'))  AND CHAR(VARCHAR_FORMAT(current date, 'YYYYMMDD')) 



SELECT  a.AAITM#, CHAR(a.AASER#) AS SN, to_date(a.AAADAT||' '||right('000000'||ltrim(a.AAATIM),6), 'yyyymmdd hh24:mi:ss') as RM_Time
FROM  DISTLIBL.ACTAUDT a
WHERE a.AAAPGM in ('HJ111E') AND ((current TIMESTAMP) - to_date(a.AAADAT||' '||right('000000'||ltrim(a.AAATIM),6), 'yyyymmdd hh24:mi:ss')) <=15



--TRX
Date('20'||Substr(T1.UPDDT, 2, 2) || '-'||  Substr(T1.UPDDT, 4, 2)|| '-' ||substr(T1.UPDDT, 6, 2)) AS "DATE",
WEEK(DATE('20'||SUBSTR(CHAR(t1.UPDDT),2,2)||'-'||SUBSTR(CHAR(t1.UPDDT),4,2)||'-'||SUBSTR(CHAR(t1.UPDDT),6,2))) AS "WEEK",
YEAR(DATE('20'||SUBSTR(CHAR(t1.UPDDT),2,2)||'-'||SUBSTR(CHAR(t1.UPDDT),4,2)||'-'||SUBSTR(CHAR(t1.UPDDT),6,2))) AS "YEAR",
(CASE 
        WHEN t2.ITCLS NOT LIKE 'Z%' THEN 'RP'
		WHEN t1.ITNBR LIKE '100-%' THEN 'CG'
		WHEN t1.ITNBR LIKE '5100-%' THEN 'CG'
		WHEN SUBSTR(t1.ITNBR,1,1) IN ('A','B','D','H','L','Q','R','T','W','Z') THEN 'CG'
		ELSE 'UPH' END) AS Product,
(CASE 
		WHEN t1.TCODE LIKE 'RP%' THEN 'Inbound'
		WHEN t1.TCODE LIKE 'SA%' THEN 'Outbound'
		ELSE 'CHECK' END) AS IN/OUT



-- between 1210901 and 1210927 
a.
T1.UPDDT BETWEEN CHAR('1'||VARCHAR_FORMAT(current date - 1 Days,'YYMMDD')) AND CHAR('1'||VARCHAR_FORMAT(current date,'YYMMDD'))

-- 将1211217 转换成周
WEEK(DATE('20'||SUBSTR(CHAR(T1.UPDDT),2,2)||'-'||SUBSTR(CHAR(T1.UPDDT),4,2)||'-'||SUBSTR(CHAR(T1.UPDDT),6,2))) AS WEEK



b.
BETWEEN  int('1'||substr(trim(char(CURRENT DATE - 30 days)),3,2)||substr(trim(char(CURRENT DATE- 30 days)),6,2)||substr(trim(char(CURRENT DATE- 30 days)),9,2)) 
AND int('1'||substr(trim(char(CURRENT DATE + 360 days)),3,2)||substr(trim(char(CURRENT DATE + 360 days)),6,2)||substr(trim(char(CURRENT DATE + 360 days)),9,2))

this is very good
--select CHAR('1'||VARCHAR_FORMAT(current date -3 days,'YYMMDD')) AS DATE from sysibm.sysdummy1
'1211101'

-- date转换为六位文本
select VARCHAR_FORMAT(current date,'YYYYMMDD') AS DATE from sysibm.sysdummy1
-- '210827'

values date(current timestamp); 
values VARCHAR_FORMAT(current TIMESTAMP,'yyyy-mm-dd');   显示 2121-08-07


-- 将 1211120 转换成日期
Date('20'||Substr(T1.UPDDT, 2, 2) || '-'||  Substr(T1.UPDDT, 4, 2)|| '-' ||substr(T1.UPDDT, 6, 2)) AS "DATE"

结果：
DATE
2021-11-20

--将文本AAADAT, AATIM 转换为日期时间 
to_date(a.AAADAT||' '||right('000000'||ltrim(a.AAATIM),6), 'yyyymmdd hh24:mi:ss') as RM_Time

/*
AAADAT	AAATIM	AAAUSR	AATARA	AATASL	AATSEC	AATTIR	AAAPGM	DATE
20211109	14502	LLQP2	RM	1	AA	1	PC228B	11/9/21 1:45
*/


(current timestamp) as CurrentTime,to_date(a.AAADAT||' '||right('000000'||ltrim(a.AAATIM),6), 'yyyymmdd hh24:mi:ss') RM_Time,
((current timestamp) - to_date(a.AAADAT||' '||right('000000'||ltrim(a.AAATIM),6), 'yyyymmdd hh24:mi:ss')) as Range,

-- 两个日期时间间隔小时
(values timestampdiff(8,char(current timestamp -  to_date(a.AAADAT||' '||right('000000'||ltrim(a.AAATIM),6), 'yyyymmdd hh24:mi:ss')))) as PendingHours


-- 定义时间区间
SELECT NDATE, int(substr(trim(char(CURRENT DATE - 10 days)),1,4)||substr(trim(char(CURRENT DATE-10 days)),6,2)||substr(trim(char(CURRENT DATE-10 days)),9,2)) AS DATE 
FROM DISTLIBQ.BTRSNCDE t1
/*
NDATE	DATE
20210517	20210612
*/


-- current date is 2021-10_9:10:54:37 提取时分秒字符
SELECT CHAR(TRIM(CHAR(HOUR(CURRENT TIMESTAMP)))||TRIM(CHAR(MINUTE(CURRENT TIMESTAMP)))||TRIM(CHAR(SECOND(CURRENT TIMESTAMP)))) FROM SYSIBM.SYSDUMMY1
--105437


-- current day是周六 10/9/2021
SELECT DAYOFWEEK(current date)    FROM sysibm.sysdummy1 
--7

-- 计算两个日期之间的天数 jimshen tested on Nov.26.2021
SELECT days(current date)- days(date('2021-11-20')) FROM sysibm.sysdummy1
-- 结果为4


-- 定义上周的开始日期与结束日期， dayofweek是本周的第几天(周日为1，2，3，4，5,6为周六), ex: current date is 10/9/2021 Sat.
select CURRENT_DATE - (DAYOFWEEK(CURRENT_DATE) + 6) DAY BEGIN_DATE,
       CURRENT_DATE - (DAYOFWEEK(CURRENT_DATE)) DAY END_DATE
FROM SYSIBM.sysdummy1
/*
BEGIN_DATE	        END_DATE
2021-09-26(周日）	2021-10-02(周六)
*/


-- 自动定位到下周六
SELECT current date + (7-DAYOFWEEK(current date)) days+7 days    FROM sysibm.sysdummy1

-- 到下下周六(10/9/2021 CURRENT DATE)
SELECT current date + (7-DAYOFWEEK(current date)) days+14 days    FROM sysibm.sysdummy1
-- 10/23/2021



--获取当前时间戳 
select current timestamp as "Data Collected At: "
 from sysibm.sysdummy1
-- 10/9/2021  7:48:51 AM
 
 
 -- 昨天白夜班的时间范围!!!!!
 between timestamp(trim(char(CURRENT DATE - 1 days))||'-19.00.00.000000') and  timestamp(trim(char(CURRENT DATE))||'-06.59.59.999999')

TIMESTAMP 的时间区间查询
Where a.WCILASTMAINTENANCETIMESTAMP  between char(current date - 21 days) and char(current DATE)  



-- time转换为六位文本
SELECT right('000000'||ltrim(X1.AAATIM),5) AS TIME    
FROM sysibm.sysdummy1 
-- 显示'000010'
 
-- date转换为六位文本
select VARCHAR_FORMAT(current date,'YYMMDD') AS DATE from sysibm.sysdummy1
-- '210827'





--获取当前日,如果今天是10/8/2021, 则显示8 
select  day(current TIMESTAMP)
from sysibm.sysdummy1
 
 
select current date from sysibm.sysdummy1; 
----------------------------------------------------------------------------------------------------------------------------------------------
-----------------------------------------------------------------------------------------------------------------------------------------------

values current date;
 
--获取当前日期 
select current time from sysibm.sysdummy1; 
values current time; 
--获取当前时间戳 
select current timestamp from sysibm.sysdummy1; 
values current timestamp; 
 
--要使当前时间或当前时间戳记调整到 GMT/CUT，则把当前的时间或时间戳记减去当前时区寄存器：
 
values current time -current timezone; 
values current timestamp -current timezone; 
 
--获取当前年份
 
values year(current timestamp);
 
--获取当前月 
values month(current timestamp);
 
--获取当前日 
select  day(current TIMESTAMP)
from sysibm.sysdummy1
 
--获取当前时 
values hour(current timestamp);
 
--获取分钟 
values minute(current timestamp);
 
--获取秒 
values second(current timestamp);
 
--获取毫秒 
values microsecond(current timestamp); 
 


values char(current date); 
values time(current timestamp); 
 
--执行日期和时间的计算
 
values current date+1 year;    
values current date+3 years+2 months +15 days; 
values current time +5 hours -3 minutes +10 seconds; 
 
--计算两个日期之间的天数
 
values days(current date)- days(date('2010-02-20')); 
 
--时间和日期换成字符串
 
values char(current date); 
values char(current time); 
 
--要将字符串转换成日期或时间值
 
values timestamp('2010-03-09-22.43.00.000000'); 
values timestamp('2010-03-09 22:44:36'); 
values date('2010-03-09'); 
values date('03/09/2010'); 
values time('22:45:27'); 
values time('22.45.27'); 
 
--计算两个时间戳记之间的时差：
 
--秒的小数部分为单位 
values timestampdiff(1,char(current timestamp - timestamp('2010-01-01-00.00.00'))); 
--秒为单位 
values timestampdiff(2,char(current timestamp - timestamp('2010-01-01-00.00.00'))); 
--分为单位 
values timestampdiff(4,char(current timestamp - timestamp('2010-01-01-00.00.00'))); 
--小时为单位 
values timestampdiff(8,char(current timestamp - timestamp('2010-01-01-00.00.00'))); 
--天为单位 
values timestampdiff(16,char(current timestamp - timestamp('2010-01-01-00.00.00'))); 
--周为单位 
values timestampdiff(32,char(current timestamp - timestamp('2010-01-01-00.00.00'))); 
--月为单位 
values timestampdiff(64,char(current timestamp - timestamp('2010-01-01-00.00.00'))); 
--季度为单位 
values timestampdiff(128,char(current timestamp - timestamp('2010-01-01-00.00.00'))); 
--年为单位 
values timestampdiff(256,char(current timestamp - timestamp('2010-01-01-00.00.00')));



--DAYNAME（）返回指定日期的星期名，该星期名是由首字符大写、其他字符小写组成的英文名。
values DAYNAME(current timestamp);--Wednesday(当天为星期五)

--DAYOFWEEK（）返回参数中的星期几，用范围在 1-7 的整数值表示，其中 1 代表星期日。
values DAYOFWEEK(current timestamp);--4(当天为星期三)

--JIMSHEN: 如下显示为2
SELECT DAYOFWEEK(current date)    FROM sysibm.sysdummy1 

-- Jimshen: 如下显示上周的开始日期与结束日期， dayofweek是本周的第几天(周日为1，2，3，4，5,6为周六)
select CURRENT_DATE - (DAYOFWEEK(CURRENT_DATE) + 5) DAY BEGIN_DATE,
       CURRENT_DATE - (DAYOFWEEK(CURRENT_DATE) - 1) DAY END_DATE
FROM SYSIBM.sysdummy1

-- 自动定位到下周六
SELECT current date + (7-DAYOFWEEK(current date)) days+7 days    FROM sysibm.sysdummy1

-- 到下下周六
SELECT current date + (7-DAYOFWEEK(current date)) days+14 days    FROM sysibm.sysdummy1




--DAYOFWEEK_ISO（）返回参数中的星期几，用范围在 1-7 的整数值表示，其中 1 代表星期一。
values DAYOFWEEK_ISO(current timestamp);--3(当前为星期三)

--DAYOFYEAR（）返回参数中一年中的第几天，用范围在 1-366 的整数值表示。
values DAYOFYEAR(current timestamp);--6

--MONTHNAME（）对于参数的月部分的月份，返回一个大小写混合的字符串（例如，January）。
values MONTHNAME(CURRENT TIMESTAMP);--January(当前为一月)

--WEEK（）返回参数中一年的第几周，用范围在 1-54 的整数值表示。以星期日作为一周的开始。（参数可以为日期格式或者日期格式的字符串）
VALUES WEEK('2016-01-02');--1
VALUES WEEK('2016-01-03');--2

--WEEK_ISO（）返回参数中一年的第几周，用范围在 1-54 的整数值表示。以星期一作为一周的开始。（参数可以为日期格式或者日期格式的字符串）
VALUES WEEK_ISO('2016-01-02');--53
VALUES WEEK_ISO('2016-01-03');--53
VALUES WEEK_ISO('2016-01-04');--1