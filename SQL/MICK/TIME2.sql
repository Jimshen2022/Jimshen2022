

--text to date

SELECT T1.TCODE,T1.ITNBR, T1.HOUSE, T2.ITCLS,
YEAR(DATE('20'||SUBSTR(CHAR(T1.UPDDT),2,2)||'-'||SUBSTR(CHAR(T1.UPDDT),4,2)||'-'||SUBSTR(CHAR(T1.UPDDT),6,2))) AS YEAR,
SUM(T1.TRQTY) TRQTY
FROM AMFLIBL.IMHIST T1, AMFLIBL.ITMRVA T2, AMFLIBL.WHSMST T3
WHERE T2.ITNBR = T1.ITNBR AND T2.STID = T3.STID AND T1.HOUSE = T3.WHID AND T1.HOUSE='51' 
AND T1.UPDDT BETWEEN '1201201' AND CHAR('1'||VARCHAR_FORMAT(current date,'YYMMDD'))
AND T1.TRQTY<>0 AND T1.TCODE IN ('RP','PQ','RM','RS','RC','IA','LA','SA','IP','IS','VR','IU','SC','SM','SP','SS') 
GROUP BY  T1.TCODE,T1.ITNBR, T1.HOUSE, T2.ITCLS, 
YEAR(DATE('20'||SUBSTR(CHAR(T1.UPDDT),2,2)||'-'||SUBSTR(CHAR(T1.UPDDT),4,2)||'-'||SUBSTR(CHAR(T1.UPDDT),6,2)))
ORDER BY T1.ITNBR 




-- 人工造时间表

with temp (date) as (
select date('01.01.2021') as date from sysibm.sysdummy1
union all
select date + 1 day from temp
where date < date('31.12.2021'))
select * from temp
	
	
/*
以上代码显示结果如下：
DATE
2/23/2016
2/24/2016
2/25/2016
2/26/2016
2/27/2016
2/28/2016
2/29/2016
3/1/2016
3/2/2016


	*/


时间加减：后边记得跟上时间类型如day、HOUR
TIMESTAMP ( TIMESTAMP(DEF_TIME)+1 day)+18 HOUR
 
DB2时间函数是我们最常见的函数之一，下面就为您介绍一些DB2时间函数，供您参考，希望可以让您对DB2时间函数有更多的了解。
--获取当前日期：   
  
select current date from sysibm.sysdummy1;    
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
values day(current timestamp);   
  
--获取当前时    
values hour(current timestamp);   
  
--获取分钟    
values minute(current timestamp);   
  
--获取秒    
values second(current timestamp);   
  
--获取毫秒    
values microsecond(current timestamp);    
  
--从时间戳记单独抽取出日期和时间   
  
values date(current timestamp);    
values VARCHAR_FORMAT(current TIMESTAMP,'yyyy-mm-dd');    
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