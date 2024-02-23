-- wanek container utlization DETAILS UPDATED ON NOVE.22.2021 
SELECT WCIORIGIN,CONTAINER#,CUBES,ITCLS,PRODUCT,WCIDESTINATION,WCIORDER,ITEMNUMBER,QTY,WCILASTMAINTENANCETIMESTAMP,CONTAINERNUMBER,DISTINCTCONTAINER,
WCILASTMAINTENANCEUSER,ITMCQTY,UNITCUBE,UNITWEIGHT,CARTONS,CONTAINERTYPE,Date,WCHCONTAINERSIZE,
(case 
when substr(WCHCONTAINERSIZE,1,1) = '4' then CUBES/2650
when substr(WCHCONTAINERSIZE,1,1) = '2' then CUBES/1325
ELSE CUBES/2650 END) AS Utilization

FROM
(SELECT s1.ContainerNumber,s1.WCIORIGIN,s1.WCIDESTINATION,s1.WCIORDER,s1.ItemNumber,s1.Qty,s1.WCILASTMAINTENANCETIMESTAMP,s1.WCILASTMAINTENANCEUSER,s1.ITMCQTY,
s1.itcls,s1.UnitCube,s1.UnitWeight,s1.Cubes,s1.Cartons,s1.Product,s2.ContainerType,to_char(s1.WCILASTMAINTENANCETIMESTAMP,'yyyy-mm-dd') as Date, 
s2.Container#,s1.Cubes/2650 as Utilization

FROM 
-- s1 union求得container当前与存档的明细
((SELECT trim(a.WCICONTAINERNUMBER) as ContainerNumber, a.WCIORIGIN, a.WCIDESTINATION, a.WCIORDER, trim(a.WCIITEMNUMBER) as ItemNumber, a.WCIQUANTITYLOADED as Qty, 
a.WCILASTMAINTENANCETIMESTAMP, a.WCILASTMAINTENANCEUSER, b.ITMCQTY, c.itcls,c.B2Z95S as UnitCube, c.WEGHT as UnitWeight, a.WCIQUANTITYLOADED*c.B2Z95S as Cubes,
CEIL(a.WCIQUANTITYLOADED/b.ITMCQTY) as Cartons,
trim(a.WCIORIGIN)||'-'|| trim(a.WCICONTAINERNUMBER)||'-'||trim(a.WCIDESTINATION) as Container#,
(CASE 
        WHEN a.WCIITEMNUMBER LIKE 'B%'  then 'CG'
        WHEN c.ITCLS not LIKE 'Z%'  then 'RP'
        WHEN c.ITCLS LIKE 'Z%' and c.ITCLS LIKE '%K' then 'Un-Kits'
        WHEN c.ITCLS LIKE 'Z%' and c.ITCLS LIKE '%Z' then 'ZipperCover'
        ELSE 'UPH' END) as Product
FROM WWUSAF.WVCNTID as a, AFILELIBW.ITMEXT as b, AMFLIBW.ITMRVA as c 
WHERE (a.WCIITEMNUMBER = b.itnbr) and a.WCIITEMNUMBER = c.itnbr and a.WCIORIGIN = c.STID and a.WCIORIGIN in('35')  
and  a.WCILASTMAINTENANCETIMESTAMP  between char(current date - 21 days) and char(current DATE)  
AND substr(trim(a.WCICONTAINERNUMBER),1,4) NOT IN ('AAAR','AIIR','AAIR','AIRR','AIR_','AIR1','AAII','ARRR')
Order by a.WCICONTAINERNUMBER, a.WCIORIGIN, a.WCILASTMAINTENANCETIMESTAMP, a.WCIDESTINATION)

union all

(SELECT trim(a.WCICONTAINERNUMBER) as ContainerNumber, a.WCIORIGIN, a.WCIDESTINATION, a.WCIORDER, trim(a.WCIITEMNUMBER) as ItemNumber, a.WCIQUANTITYLOADED as Qty, 
a.WCILASTMAINTENANCETIMESTAMP, a.WCILASTMAINTENANCEUSER, b.ITMCQTY, c.itcls,c.B2Z95S as UnitCube, c.WEGHT as UnitWeight, a.WCIQUANTITYLOADED*c.B2Z95S as Cubes,
CEIL(a.WCIQUANTITYLOADED/b.ITMCQTY) as Cartons,
trim(a.WCIORIGIN)||'-'|| trim(a.WCICONTAINERNUMBER)||'-'||trim(a.WCIDESTINATION)||'-'||SUBSTR(char(a.WCIARCHIVETIMESTAMP),1,13) as Container#,
(CASE 
        WHEN a.WCIITEMNUMBER LIKE 'B%'  then 'CG'
        WHEN c.ITCLS not LIKE 'Z%'  then 'RP'
        WHEN c.ITCLS LIKE 'Z%' and c.ITCLS LIKE '%K' then 'Un-Kits'
        WHEN c.ITCLS LIKE 'Z%' and c.ITCLS LIKE '%Z' then 'ZipperCover'
        ELSE 'UPH' END) as Product
FROM ASHLEYARCW.WVCNTIDA as a, AFILELIBW.ITMEXT as b, AMFLIBW.ITMRVA as c
WHERE (a.WCIITEMNUMBER = b.itnbr) and a.WCIITEMNUMBER = c.itnbr and a.WCIORIGIN = c.STID and a.WCIORIGIN in('35') and a.WCILASTMAINTENANCETIMESTAMP  
between char(current date - 21 days) and char(current DATE)  AND substr(a.WCICONTAINERNUMBER,1,4) NOT IN ('AAAR','AIIR','AAIR','AIRR','AIR_','AIR1','AAII','ARRR')
Order by a.WCICONTAINERNUMBER, a.WCIORIGIN, a.WCILASTMAINTENANCETIMESTAMP, a.WCIDESTINATION)) as s1,

-- s2 求得container是否为混装或非混装
(SELECT Container#,(case when count(Distinct PRODUCT)=1 then 'None-Mixed'  else 'Mixed' end) as ContainerType
FROM 
((SELECT trim(a.WCICONTAINERNUMBER) as ContainerNumber, a.WCIORIGIN, a.WCIDESTINATION, a.WCIORDER, trim(a.WCIITEMNUMBER) as ItemNumber, a.WCIQUANTITYLOADED as Qty, 
a.WCILASTMAINTENANCETIMESTAMP, a.WCILASTMAINTENANCEUSER, b.ITMCQTY, c.itcls,c.B2Z95S as UnitCube, c.WEGHT as UnitWeight, a.WCIQUANTITYLOADED*c.B2Z95S as Cubes,
CEIL(a.WCIQUANTITYLOADED/b.ITMCQTY) as Cartons,
trim(a.WCIORIGIN)||'-'|| trim(a.WCICONTAINERNUMBER)||'-'||trim(a.WCIDESTINATION) as Container#,
(CASE 
        WHEN a.WCIITEMNUMBER LIKE 'B%'  then 'CG'
        WHEN c.ITCLS not LIKE 'Z%'  then 'RP'
        WHEN c.ITCLS LIKE 'Z%' and c.ITCLS LIKE '%K' then 'Un-Kits'
        WHEN c.ITCLS LIKE 'Z%' and c.ITCLS LIKE '%Z' then 'ZipperCover'
        ELSE 'UPH' END) as Product
FROM WWUSAF.WVCNTID as a, AFILELIBW.ITMEXT as b, AMFLIBW.ITMRVA as c
WHERE (a.WCIITEMNUMBER = b.itnbr) and a.WCIITEMNUMBER = c.itnbr and a.WCIORIGIN = c.STID and a.WCIORIGIN in('35')  
and  a.WCILASTMAINTENANCETIMESTAMP BETWEEN char(current date - 21 days) and char(current DATE) 
AND substr(trim(a.WCICONTAINERNUMBER),1,4) NOT IN ('AAAR','AIIR','AAIR','AIRR','AIR_','AIR1')
Order by a.WCICONTAINERNUMBER, a.WCIORIGIN, a.WCILASTMAINTENANCETIMESTAMP, a.WCIDESTINATION)

union all

(SELECT trim(a.WCICONTAINERNUMBER) as ContainerNumber, a.WCIORIGIN, a.WCIDESTINATION, a.WCIORDER, trim(a.WCIITEMNUMBER) as ItemNumber, a.WCIQUANTITYLOADED as Qty, 
a.WCILASTMAINTENANCETIMESTAMP, a.WCILASTMAINTENANCEUSER, b.ITMCQTY, c.itcls,c.B2Z95S as UnitCube, c.WEGHT as UnitWeight, a.WCIQUANTITYLOADED*c.B2Z95S as Cubes,
CEIL(a.WCIQUANTITYLOADED/b.ITMCQTY) as Cartons,
trim(a.WCIORIGIN)||'-'|| trim(a.WCICONTAINERNUMBER)||'-'||trim(a.WCIDESTINATION)||'-'||SUBSTR(char(a.WCIARCHIVETIMESTAMP),1,13) as Container#,
(CASE 
        WHEN a.WCIITEMNUMBER LIKE 'B%'  then 'CG'
        WHEN c.ITCLS not LIKE 'Z%'  then 'RP'
        WHEN c.ITCLS LIKE 'Z%' and c.ITCLS LIKE '%K' then 'Un-Kits'
        WHEN c.ITCLS LIKE 'Z%' and c.ITCLS LIKE '%Z' then 'ZipperCover'
        ELSE 'UPH' END) as Product
FROM ASHLEYARCW.WVCNTIDA as a, AFILELIBW.ITMEXT as b, AMFLIBW.ITMRVA as c
WHERE (a.WCIITEMNUMBER = b.itnbr) and a.WCIITEMNUMBER = c.itnbr and a.WCIORIGIN = c.STID and a.WCIORIGIN in('35') and a.WCILASTMAINTENANCETIMESTAMP 
between char(current date - 21 days) and char(current DATE) AND substr(a.WCICONTAINERNUMBER,1,4) NOT IN ('AAAR','AIIR','AAIR','AIRR','AIR_','AIR1','AAII','ARRR')
Order by a.WCICONTAINERNUMBER, a.WCIORIGIN, a.WCILASTMAINTENANCETIMESTAMP, a.WCIDESTINATION))
group by WCIORIGIN, Container#
order by Container#) as s2

WHERE s1.Container# = s2.Container#
order by s1.WCIORIGIN,s1.ContainerNumber,s1.WCILASTMAINTENANCETIMESTAMP
) AS Y1

right join
 -- rigt join 目的是找出Header文件中的集装箱号与detail匹配
(Select DISTINCT(x1.Container#) as DistinctContainer, WCHCONTAINERSIZE

FROM
(SELECT 
a.WCHCONTAINERNUMBER,a.WCHORIGIN,a.WCHDESTINATION,a.WCHCONTAINERSTATUS,a.WCHTOTALCARTONS,a.WCHTOTALCUBES,a.WCHPOSTEDTIMESTAMP,a.WCHTOTALWEIGHT,
 a.WCHCONTAINERSIZE,
 trim(a.WCHORIGIN)||'-'|| trim(a.WCHCONTAINERNUMBER)||'-'||trim(a.WCHDESTINATION) as Container#
FROM  WWUSAF.WVCNTHD a
WHERE a.WCHCONTAINERSTATUS in ('P','T') AND a.WCHORIGIN IN ('35') and (a.WCHACTUALARRIVALMAINTPROGRAM='SVCHECKIN' 
OR (a.WCHACTUALARRIVALMAINTPROGRAM not in ('SVCHECKIN') and WCHBUILDING in ('B1','V3','M3')))
AND a.WCHPOSTEDTIMESTAMP BETWEEN char(current date - 14 days) and char(current DATE) 
and substr(trim(a.WCHCONTAINERNUMBER),1,4) NOT IN ('AAAR','AIIR','AAIR','AIRR','AIR_','AIR1','AAII','ARRR')

union all

SELECT  a.WCHCONTAINERNUMBER,a.WCHORIGIN,a.WCHDESTINATION,a.WCHCONTAINERSTATUS,a.WCHTOTALCARTONS,a.WCHTOTALCUBES,a.WCHPOSTEDTIMESTAMP,a.WCHTOTALWEIGHT,a.WCHCONTAINERSIZE,
trim(a.WCHORIGIN)||'-'|| trim(a.WCHCONTAINERNUMBER)||'-'||trim(a.WCHDESTINATION)||'-'||SUBSTR(char(a.WCHARCHIVETIMESTAMP),1,13) as Container#
FROM  ASHLEYARCW.WVCNTHDA a
WHERE a.WCHCONTAINERSTATUS in ('P','T') and (a.WCHACTUALARRIVALMAINTPROGRAM='SVCHECKIN'
OR (a.WCHACTUALARRIVALMAINTPROGRAM not in ('SVCHECKIN') and WCHBUILDING in ('B1','V3','M3')))
AND a.WCHPOSTEDTIMESTAMP BETWEEN char(current date - 14 days) and char(current DATE)  and a.WCHORIGIN in ('35') and 
substr(trim(a.WCHCONTAINERNUMBER),1,4) NOT IN ('AAAR','AIIR','AAIR','AIRR','AIR_','AIR1','AAII','ARRR')) as x1) as Y2


ON Y1.CONTAINER# = Y2.DistinctContainer