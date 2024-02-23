-- Updated on Oct.12.2021 for wanek Container Utilization Summary sheet

Select 
x1.WCHDOORNUMBER,x1.WCHCONTAINERNUMBER,x1.WCHORIGIN,x1.WCHDESTINATION,x1.WCHCONTAINERSTATUS,x1.WCHTOTALCARTONS,x1.WCHTOTALCUBES,x1.WCHPOSTEDTIMESTAMP,x1.WCHTOTALWEIGHT,x1.WCHCONTAINERSIZE,
x1.Container#,to_char(x1.WCHPOSTEDTIMESTAMP,'yyyy-mm-dd') as Date, x1.WCHPOSTEDTIMESTAMP,x1.WCHPOSTEDUSER,
(case 
when substr(x1.WCHCONTAINERSIZE,1,1) = '4' then x1.WCHTOTALCUBES/2650
when substr(x1.WCHCONTAINERSIZE,1,1) = '2' then x1.WCHTOTALCUBES/1325
ELSE x1.WCHTOTALCUBES/2650 END) AS Utilization,
(CASE 
         WHEN CONTAINER# LIKE ('31%') THEN 'WN1'
         WHEN CONTAINER# LIKE ('33%') THEN 'WN2'
         WHEN  CONTAINER# LIKE ('35%') and TRIM(WCHDOORNUMBER) LIKE ('4%') THEN 'WN3'
         WHEN  CONTAINER# LIKE ('35%') and TRIM(WCHDOORNUMBER) LIKE ('9%') THEN 'BW'
         WHEN  CONTAINER# LIKE ('35%') and TRIM(WCHDOORNUMBER) LIKE ('8%') THEN 'DC'
         WHEN  CONTAINER# LIKE ('35%') and CONTAINER# LIKE ('%-335%') and TRIM(WCHDOORNUMBER) LIKE ('0%') THEN 'DC'
         WHEN  CONTAINER# LIKE ('35%') and CONTAINER# LIKE ('%-CNW%') and TRIM(WCHDOORNUMBER) LIKE ('0%') THEN 'DC'
         WHEN  CONTAINER# LIKE ('35%') and CONTAINER# LIKE ('%-C') and TRIM(WCHDOORNUMBER) LIKE ('0%') THEN 'DC'
         WHEN  CONTAINER# LIKE ('35%') and CONTAINER# NOT LIKE ('%-335%') and TRIM(WCHDOORNUMBER) LIKE ('0%') THEN 'WN3'
         WHEN  CONTAINER# LIKE ('35%') and CONTAINER# NOT LIKE ('%-CNW%') and TRIM(WCHDOORNUMBER) LIKE ('0%') THEN 'WN3'
         WHEN  CONTAINER# LIKE ('35%') and CONTAINER# NOT LIKE ('%-C') and TRIM(WCHDOORNUMBER) LIKE ('0%') THEN 'WN3'
 ELSE 'WN3' END) as Site

FROM

(SELECT 
a.WCHDOORNUMBER,a.WCHCONTAINERNUMBER,a.WCHORIGIN,a.WCHDESTINATION,a.WCHCONTAINERSTATUS,a.WCHTOTALCARTONS,a.WCHTOTALCUBES,a.WCHPOSTEDTIMESTAMP,a.WCHTOTALWEIGHT,a.WCHPOSTEDUSER,a.WCHCONTAINERSIZE,
trim(a.WCHORIGIN)||'-'|| trim(a.WCHCONTAINERNUMBER)||'-'||trim(a.WCHDESTINATION) as Container#
FROM  WWUSAF.WVCNTHD a
WHERE a.WCHCONTAINERSTATUS in ('P','T') AND a.WCHORIGIN IN ('35')  and (a.WCHACTUALARRIVALMAINTPROGRAM='SVCHECKIN' 
OR (a.WCHACTUALARRIVALMAINTPROGRAM not in ('SVCHECKIN') and WCHBUILDING in ('B1','V3','M3')))
AND a.WCHPOSTEDTIMESTAMP BETWEEN char(current date - 14 days) and char(current DATE) 
and substr(trim(a.WCHCONTAINERNUMBER),1,4) NOT IN ('AAAR','AIIR','AAIR','AIRR','AIR_','AIR1')

union all

SELECT  a.WCHDOORNUMBER,a.WCHCONTAINERNUMBER,a.WCHORIGIN,a.WCHDESTINATION,a.WCHCONTAINERSTATUS,a.WCHTOTALCARTONS,a.WCHTOTALCUBES,a.WCHPOSTEDTIMESTAMP,a.WCHTOTALWEIGHT,a.WCHPOSTEDUSER,a.WCHCONTAINERSIZE,
trim(a.WCHORIGIN)||'-'|| trim(a.WCHCONTAINERNUMBER)||'-'||trim(a.WCHDESTINATION) as Container#
FROM  ASHLEYARCW.WVCNTHDA a
WHERE a.WCHCONTAINERSTATUS in ('P','T') and (a.WCHACTUALARRIVALMAINTPROGRAM='SVCHECKIN' 
OR (a.WCHACTUALARRIVALMAINTPROGRAM not in ('SVCHECKIN') and WCHBUILDING in ('B1','V3','M3')))
AND a.WCHPOSTEDTIMESTAMP BETWEEN char(current date - 14 days) and char(current DATE)  and a.WCHORIGIN in ('35') and 
substr(trim(a.WCHCONTAINERNUMBER),1,4) NOT IN ('AAAR','AIIR','AAIR','AIRR','AIR_','AIR1')) as x1