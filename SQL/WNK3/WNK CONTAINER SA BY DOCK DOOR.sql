-- WANEK CONTAINER SA BY DOCK DOOR AND HOUR, CREATED BY JIMSHEN ON Oct.5.2022 
--details
(SELECT 
trim(a.WCHORIGIN)||'-'|| trim(a.WCHCONTAINERNUMBER)||'-'||trim(a.WCHDESTINATION) as Container#,
a.WCHCONTAINERNUMBER,a.WCHORIGIN,a.WCHDESTINATION,a.WCHCONTAINERSTATUS,a.WCHTOTALCARTONS,a.WCHTOTALCUBES,a.WCHPOSTEDTIMESTAMP,a.WCHTOTALWEIGHT,
 a.WCHCONTAINERSIZE,a.WCHDOORNUMBER,a.WCHDOCKID,a.WCHBUILDING,a.WCHSEALNUMBER,a.WCHCARRIER,char(hour(a.WCHPOSTEDTIMESTAMP)) as hour

FROM  DISTLIBW.TBL_WVCONTAINER_HDR a
WHERE a.WCHCONTAINERSTATUS in ('P','T') AND a.WCHORIGIN IN ('35') and (a.WCHACTUALARRIVALMAINTPROGRAM='SVCHECKIN' 
OR (a.WCHACTUALARRIVALMAINTPROGRAM not in ('SVCHECKIN') and WCHBUILDING in ('B1','V3','M3')))
AND a.WCHCONTAINERNUMBER not like 'AIR%' AND a.WCHDESTINATION NOT IN ('100','101','102','12','131','01','3','990')
AND a.WCHPOSTEDTIMESTAMP BETWEEN ? and ?
and substr(trim(a.WCHCONTAINERNUMBER),1,4) NOT IN ('AAAR','AIIR','AAIR','AIRR','AIR_','AIR1','AAII','ARRR')

union all

SELECT 
trim(a.WCHORIGIN)||'-'|| trim(a.WCHCONTAINERNUMBER)||'-'||trim(a.WCHDESTINATION)||'-'||SUBSTR(char(a.WCHARCHIVETIMESTAMP),1,13) as Container#,
 a.WCHCONTAINERNUMBER,a.WCHORIGIN,a.WCHDESTINATION,a.WCHCONTAINERSTATUS,a.WCHTOTALCARTONS,a.WCHTOTALCUBES,a.WCHPOSTEDTIMESTAMP,a.WCHTOTALWEIGHT,a.WCHCONTAINERSIZE,a.WCHDOORNUMBER,a.WCHDOCKID,a.WCHBUILDING,a.WCHSEALNUMBER,a.WCHCARRIER,char(hour(a.WCHPOSTEDTIMESTAMP)) as hour

FROM  ASHLEYARCW.TBL_WVCONTAINER_HDR_A a
WHERE a.WCHCONTAINERSTATUS in ('P','T') and (a.WCHACTUALARRIVALMAINTPROGRAM='SVCHECKIN'
OR (a.WCHACTUALARRIVALMAINTPROGRAM not in ('SVCHECKIN') and WCHBUILDING in ('B1','V3','M3')))
AND a.WCHCONTAINERNUMBER not like 'AIR%' AND a.WCHDESTINATION NOT IN ('100','101','102','12','131','01','3','990')
AND a.WCHPOSTEDTIMESTAMP BETWEEN ? and ?  and a.WCHORIGIN in ('35') and 
substr(trim(a.WCHCONTAINERNUMBER),1,4) NOT IN ('AAAR','AIIR','AAIR','AIRR','AIR_','AIR1','AAII','ARRR')) as x1



--summary

(SELECT 
 trim(a.WCHORIGIN)||'-'|| trim(a.WCHCONTAINERNUMBER)||'-'||trim(a.WCHDESTINATION) as Container#,
a.WCHCONTAINERNUMBER,a.WCHORIGIN,a.WCHDESTINATION,a.WCHCONTAINERSTATUS,a.WCHTOTALCARTONS,a.WCHTOTALCUBES,a.WCHPOSTEDTIMESTAMP,a.WCHTOTALWEIGHT,
 a.WCHCONTAINERSIZE,a.WCHDOORNUMBER,a.WCHDOCKID,a.WCHBUILDING,a.WCHSEALNUMBER,a.WCHCARRIER,char(hour(a.WCHPOSTEDTIMESTAMP)) hour

FROM  DISTLIBW.TBL_WVCONTAINER_HDR a
WHERE a.WCHCONTAINERSTATUS in ('P','T') AND a.WCHORIGIN IN ('35') and (a.WCHACTUALARRIVALMAINTPROGRAM='SVCHECKIN' 
OR (a.WCHACTUALARRIVALMAINTPROGRAM not in ('SVCHECKIN') and WCHBUILDING in ('B1','V3','M3')))
AND a.WCHCONTAINERNUMBER not like 'AIR%' AND a.WCHDESTINATION NOT IN ('100','101','102','12','131','01','3','990')
AND a.WCHPOSTEDTIMESTAMP BETWEEN ? and ?
and substr(trim(a.WCHCONTAINERNUMBER),1,4) NOT IN ('AAAR','AIIR','AAIR','AIRR','AIR_','AIR1','AAII','ARRR')

union all

SELECT  a.WCHCONTAINERNUMBER,a.WCHORIGIN,a.WCHDESTINATION,a.WCHCONTAINERSTATUS,a.WCHTOTALCARTONS,a.WCHTOTALCUBES,a.WCHPOSTEDTIMESTAMP,a.WCHTOTALWEIGHT,a.WCHCONTAINERSIZE,a.WCHDOORNUMBER,a.WCHDOCKID,a.WCHBUILDING,a.WCHSEALNUMBER,a.WCHCARRIER,
trim(a.WCHORIGIN)||'-'|| trim(a.WCHCONTAINERNUMBER)||'-'||trim(a.WCHDESTINATION)||'-'||SUBSTR(char(a.WCHARCHIVETIMESTAMP),1,13) as Container#
FROM  ASHLEYARCW.TBL_WVCONTAINER_HDR_A a
WHERE a.WCHCONTAINERSTATUS in ('P','T') and (a.WCHACTUALARRIVALMAINTPROGRAM='SVCHECKIN'
OR (a.WCHACTUALARRIVALMAINTPROGRAM not in ('SVCHECKIN') and WCHBUILDING in ('B1','V3','M3')))
AND a.WCHCONTAINERNUMBER not like 'AIR%' AND a.WCHDESTINATION NOT IN ('100','101','102','12','131','01','3','990')
AND a.WCHPOSTEDTIMESTAMP BETWEEN ? and ?  and a.WCHORIGIN in ('35') and 
substr(trim(a.WCHCONTAINERNUMBER),1,4) NOT IN ('AAAR','AIIR','AAIR','AIRR','AIR_','AIR1','AAII','ARRR')) as x1