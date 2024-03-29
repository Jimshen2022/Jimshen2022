-- Updated on Nov.23.2021 for wanek Container Utilization Summary sheet

SELECT 
x1.WCHDOORNUMBER,x1.WCHCONTAINERNUMBER,x1.WCHORIGIN,x1.WCHDESTINATION,x1.WCHCONTAINERSTATUS,x1.WCHTOTALCARTONS,x1.WCHTOTALCUBES,x1.WCHPOSTEDTIMESTAMP,x1.WCHTOTALWEIGHT,x1.WCHCONTAINERSIZE,
x1.Container#,to_char(x1.WCHPOSTEDTIMESTAMP,'yyyy-mm-dd') as Date, x1.WCHPOSTEDTIMESTAMP,x1.WCHPOSTEDUSER,
(CASE 
	WHEN TRIM(SUBSTR(x1.WCHCONTAINERSIZE,1,3)) = '40H' THEN x1.WCHTOTALCUBES/2650
	WHEN TRIM(SUBSTR(x1.WCHCONTAINERSIZE,1,3)) = '40' THEN x1.WCHTOTALCUBES/2383
	WHEN TRIM(SUBSTR(x1.WCHCONTAINERSIZE,1,3)) = '45' THEN x1.WCHTOTALCUBES/3058
	WHEN SUBSTR(x1.WCHCONTAINERSIZE,1,1) = '2' THEN x1.WCHTOTALCUBES/1191
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
WHERE a.WCHCONTAINERSTATUS in ('P','T') AND a.WCHORIGIN IN ('33')  and (a.WCHACTUALARRIVALMAINTPROGRAM='SVCHECKIN' 
OR (a.WCHACTUALARRIVALMAINTPROGRAM not in ('SVCHECKIN') and WCHBUILDING in ('33'))) and a.WCHCONTAINERNUMBER not like 'SUNR2%'  
AND a.WCHCONTAINERNUMBER not like 'AIR%' AND a.WCHDESTINATION NOT IN ('100','101','102','12','131','01','3','990') AND a.WCHBUILDING NOT IN ('B2','B4','B5')
AND a.WCHPOSTEDTIMESTAMP BETWEEN char(current date - 14 days) and char(current DATE) 
and substr(trim(a.WCHCONTAINERNUMBER),1,4) NOT IN ('AAAR','AIIR','AAIR','AIRR','AIR_','AIR1')

union all

SELECT  a.WCHDOORNUMBER,a.WCHCONTAINERNUMBER,a.WCHORIGIN,a.WCHDESTINATION,a.WCHCONTAINERSTATUS,a.WCHTOTALCARTONS,a.WCHTOTALCUBES,a.WCHPOSTEDTIMESTAMP,a.WCHTOTALWEIGHT,a.WCHPOSTEDUSER,a.WCHCONTAINERSIZE,
trim(a.WCHORIGIN)||'-'|| trim(a.WCHCONTAINERNUMBER)||'-'||trim(a.WCHDESTINATION)||'-'||SUBSTR(char(a.WCHARCHIVETIMESTAMP),1,13) as Container#
FROM  ASHLEYARCW.WVCNTHDA a
WHERE a.WCHCONTAINERSTATUS in ('P','T') and (a.WCHACTUALARRIVALMAINTPROGRAM='SVCHECKIN' 
OR (a.WCHACTUALARRIVALMAINTPROGRAM not in ('SVCHECKIN') and WCHBUILDING in ('33'))) and a.WCHCONTAINERNUMBER not like 'SUNR2%' 
AND a.WCHCONTAINERNUMBER not like 'AIR%'  AND a.WCHDESTINATION NOT IN ('100','101','102','12','131','01','3','990') AND a.WCHBUILDING NOT IN ('B2','B4','B5')
AND a.WCHPOSTEDTIMESTAMP BETWEEN char(current date - 14 days) and char(current DATE)  and a.WCHORIGIN in ('33') and 
substr(trim(a.WCHCONTAINERNUMBER),1,4) NOT IN ('AAAR','AIIR','AAIR','AIRR','AIR_','AIR1')) as x1