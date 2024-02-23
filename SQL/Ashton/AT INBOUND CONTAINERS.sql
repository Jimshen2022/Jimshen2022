


SELECT  a.AACOD1,a.AACOD3,a.AAEQP#, a.AAADAT,a.AATFR#,
week(date(substr(trim(char(a.AAADAT)),1,4)||'-'||substr(trim(char(a.AAADAT)),5,2)||'-'||substr(trim(char(a.AAADAT)),7,2))) YWeek,
COUNT(a.AAITM#) as SKUs,SUM(a.AATQTY) Qty

FROM  DISTLIB.ACTAUDT a

WHERE a.AACOD1 in ('RC') AND a.AATWHS in ('335') and a.AACOD3 in ('YD') and a.AAADAT BETWEEN '20210425' and '20211111'
Group by a.AACOD1,a.AACOD3,a.AAEQP#, a.AAADAT,a.AATFR#,
week(date(substr(trim(char(a.AAADAT)),1,4)||'-'||substr(trim(char(a.AAADAT)),5,2)||'-'||substr(trim(char(a.AAADAT)),7,2)))