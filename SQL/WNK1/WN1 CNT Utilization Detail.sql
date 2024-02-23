SELECT *
FROM
(SELECT s1.ContainerNumber,s1.WCIORIGIN,s1.WCIDESTINATION,s1.WCIORDER,s1.ItemNumber,s1.Qty,s1.WCILASTMAINTENANCETIMESTAMP,s1.WCILASTMAINTENANCEUSER,s1.ITMCQTY,
s1.itcls,s1.UnitCube,s1.UnitWeight,s1.Cubes,s1.Cartons,s1.Product,s2.ContainerType,to_char(s1.WCILASTMAINTENANCETIMESTAMP,'yyyy-mm-dd') as Date, 
s2.Container#,s1.Cubes/2650 as Utilization
FROM 
((SELECT trim(a.WCICONTAINERNUMBER) as ContainerNumber, a.WCIORIGIN, a.WCIDESTINATION, a.WCIORDER, trim(a.WCIITEMNUMBER) as ItemNumber, a.WCIQUANTITYLOADED as Qty, 
a.WCILASTMAINTENANCETIMESTAMP, a.WCILASTMAINTENANCEUSER, b.ITMCQTY, c.itcls,c.B2Z95S as UnitCube, c.WEGHT as UnitWeight, a.WCIQUANTITYLOADED*c.B2Z95S as Cubes,
CEIL(a.WCIQUANTITYLOADED/b.ITMCQTY) as Cartons,
(CASE 
        When a.WCIDESTINATION in ('335','CNW') then trim(a.WCIORIGIN)||'-'|| trim(a.WCICONTAINERNUMBER)||'-'||trim(a.WCIDESTINATION)||'-'||to_char(a.WCILASTMAINTENANCETIMESTAMP,'yyyy-mm-dd') 
        ELSE trim(a.WCIORIGIN)||'-'|| trim(a.WCICONTAINERNUMBER)||'-'||trim(a.WCIDESTINATION) END) as Container#,
(CASE 
        WHEN a.WCIITEMNUMBER LIKE 'B%'  then 'CG'
        WHEN c.ITCLS not LIKE 'Z%'  then 'RP'
        WHEN c.ITCLS LIKE 'Z%' and c.ITCLS LIKE '%K' then 'Un-Kits'
        WHEN c.ITCLS LIKE 'Z%' and c.ITCLS LIKE '%Z' then 'ZipperCover'
        ELSE 'UPH' END) as Product
FROM WWUSAF.WVCNTID as a, AFILELIBW.ITMEXT as b, AMFLIBW.ITMRVA as c 
WHERE (a.WCIITEMNUMBER = b.itnbr) and a.WCIITEMNUMBER = c.itnbr and a.WCIORIGIN = c.STID and a.WCIORIGIN in('31')  
and  a.WCILASTMAINTENANCETIMESTAMP  between char(current date - 21 days) and char(current DATE)  
AND substr(trim(a.WCICONTAINERNUMBER),1,4) NOT IN ('AAAR','AIIR','AAIR','AIRR','AIR_','AIR1')
Order by a.WCICONTAINERNUMBER, a.WCIORIGIN, a.WCILASTMAINTENANCETIMESTAMP, a.WCIDESTINATION)

union all

(SELECT trim(a.WCICONTAINERNUMBER) as ContainerNumber, a.WCIORIGIN, a.WCIDESTINATION, a.WCIORDER, trim(a.WCIITEMNUMBER) as ItemNumber, a.WCIQUANTITYLOADED as Qty, 
a.WCILASTMAINTENANCETIMESTAMP, a.WCILASTMAINTENANCEUSER, b.ITMCQTY, c.itcls,c.B2Z95S as UnitCube, c.WEGHT as UnitWeight, a.WCIQUANTITYLOADED*c.B2Z95S as Cubes,
CEIL(a.WCIQUANTITYLOADED/b.ITMCQTY) as Cartons,
(CASE 
        When a.WCIDESTINATION in ('335','CNW') then trim(a.WCIORIGIN)||'-'|| trim(a.WCICONTAINERNUMBER)||'-'||trim(a.WCIDESTINATION)||'-'||to_char(a.WCILASTMAINTENANCETIMESTAMP,'yyyy-mm-dd') 
        ELSE trim(a.WCIORIGIN)||'-'|| trim(a.WCICONTAINERNUMBER)||'-'||trim(a.WCIDESTINATION) END) as Container#,
(CASE 
        WHEN a.WCIITEMNUMBER LIKE 'B%'  then 'CG'
        WHEN c.ITCLS not LIKE 'Z%'  then 'RP'
        WHEN c.ITCLS LIKE 'Z%' and c.ITCLS LIKE '%K' then 'Un-Kits'
        WHEN c.ITCLS LIKE 'Z%' and c.ITCLS LIKE '%Z' then 'ZipperCover'
        ELSE 'UPH' END) as Product
FROM ASHLEYARCW.WVCNTIDA as a, AFILELIBW.ITMEXT as b, AMFLIBW.ITMRVA as c
WHERE (a.WCIITEMNUMBER = b.itnbr) and a.WCIITEMNUMBER = c.itnbr and a.WCIORIGIN = c.STID and a.WCIORIGIN in('31') and a.WCILASTMAINTENANCETIMESTAMP  
between char(current date - 21 days) and char(current DATE)  AND substr(a.WCICONTAINERNUMBER,1,4) NOT IN ('AAAR','AIIR','AAIR','AIRR','AIR_','AIR1')
Order by a.WCICONTAINERNUMBER, a.WCIORIGIN, a.WCILASTMAINTENANCETIMESTAMP, a.WCIDESTINATION)) as s1,


(SELECT Container#,(case when count(Distinct PRODUCT)=1 then 'None-Mixed'  else 'Mixed' end) as ContainerType
FROM 
((SELECT trim(a.WCICONTAINERNUMBER) as ContainerNumber, a.WCIORIGIN, a.WCIDESTINATION, a.WCIORDER, trim(a.WCIITEMNUMBER) as ItemNumber, a.WCIQUANTITYLOADED as Qty, 
a.WCILASTMAINTENANCETIMESTAMP, a.WCILASTMAINTENANCEUSER, b.ITMCQTY, c.itcls,c.B2Z95S as UnitCube, c.WEGHT as UnitWeight, a.WCIQUANTITYLOADED*c.B2Z95S as Cubes,
CEIL(a.WCIQUANTITYLOADED/b.ITMCQTY) as Cartons,
(CASE 
        When a.WCIDESTINATION in ('335','CNW') then trim(a.WCIORIGIN)||'-'|| trim(a.WCICONTAINERNUMBER)||'-'||trim(a.WCIDESTINATION)||'-'||to_char(a.WCILASTMAINTENANCETIMESTAMP,'yyyy-mm-dd') 
        ELSE trim(a.WCIORIGIN)||'-'|| trim(a.WCICONTAINERNUMBER)||'-'||trim(a.WCIDESTINATION) END) as Container#,
(CASE 
        WHEN a.WCIITEMNUMBER LIKE 'B%'  then 'CG'
        WHEN c.ITCLS not LIKE 'Z%'  then 'RP'
        WHEN c.ITCLS LIKE 'Z%' and c.ITCLS LIKE '%K' then 'Un-Kits'
        WHEN c.ITCLS LIKE 'Z%' and c.ITCLS LIKE '%Z' then 'ZipperCover'
        ELSE 'UPH' END) as Product
FROM WWUSAF.WVCNTID as a, AFILELIBW.ITMEXT as b, AMFLIBW.ITMRVA as c
WHERE (a.WCIITEMNUMBER = b.itnbr) and a.WCIITEMNUMBER = c.itnbr and a.WCIORIGIN = c.STID and a.WCIORIGIN in('31')  
and  a.WCILASTMAINTENANCETIMESTAMP BETWEEN char(current date - 21 days) and char(current DATE) 
AND substr(trim(a.WCICONTAINERNUMBER),1,4) NOT IN ('AAAR','AIIR','AAIR','AIRR','AIR_','AIR1')
Order by a.WCICONTAINERNUMBER, a.WCIORIGIN, a.WCILASTMAINTENANCETIMESTAMP, a.WCIDESTINATION)

union all

(SELECT trim(a.WCICONTAINERNUMBER) as ContainerNumber, a.WCIORIGIN, a.WCIDESTINATION, a.WCIORDER, trim(a.WCIITEMNUMBER) as ItemNumber, a.WCIQUANTITYLOADED as Qty, 
a.WCILASTMAINTENANCETIMESTAMP, a.WCILASTMAINTENANCEUSER, b.ITMCQTY, c.itcls,c.B2Z95S as UnitCube, c.WEGHT as UnitWeight, a.WCIQUANTITYLOADED*c.B2Z95S as Cubes,
CEIL(a.WCIQUANTITYLOADED/b.ITMCQTY) as Cartons,
(CASE 
        When a.WCIDESTINATION in ('335','CNW') then trim(a.WCIORIGIN)||'-'|| trim(a.WCICONTAINERNUMBER)||'-'||trim(a.WCIDESTINATION)||'-'||to_char(a.WCILASTMAINTENANCETIMESTAMP,'yyyy-mm-dd') 
        ELSE trim(a.WCIORIGIN)||'-'|| trim(a.WCICONTAINERNUMBER)||'-'||trim(a.WCIDESTINATION) END) as Container#,
(CASE 
        WHEN a.WCIITEMNUMBER LIKE 'B%'  then 'CG'
        WHEN c.ITCLS not LIKE 'Z%'  then 'RP'
        WHEN c.ITCLS LIKE 'Z%' and c.ITCLS LIKE '%K' then 'Un-Kits'
        WHEN c.ITCLS LIKE 'Z%' and c.ITCLS LIKE '%Z' then 'ZipperCover'
        ELSE 'UPH' END) as Product
FROM ASHLEYARCW.WVCNTIDA as a, AFILELIBW.ITMEXT as b, AMFLIBW.ITMRVA as c
WHERE (a.WCIITEMNUMBER = b.itnbr) and a.WCIITEMNUMBER = c.itnbr and a.WCIORIGIN = c.STID and a.WCIORIGIN in('31') and a.WCILASTMAINTENANCETIMESTAMP 
between char(current date - 21 days) and char(current DATE) AND substr(a.WCICONTAINERNUMBER,1,4) NOT IN ('AAAR','AIIR','AAIR','AIRR','AIR_','AIR1')
Order by a.WCICONTAINERNUMBER, a.WCIORIGIN, a.WCILASTMAINTENANCETIMESTAMP, a.WCIDESTINATION))
group by WCIORIGIN, Container#
order by Container#) as s2

WHERE s1.Container# = s2.Container#
order by s1.WCIORIGIN,s1.ContainerNumber,s1.WCILASTMAINTENANCETIMESTAMP
) AS Y1

inner join

(Select DISTINCT(x1.WCHCONTAINERNUMBER) as DistinctContainer

FROM
(SELECT 
a.WCHDOORNUMBER,a.WCHCONTAINERNUMBER,a.WCHORIGIN,a.WCHDESTINATION,a.WCHCONTAINERSTATUS,a.WCHTOTALCARTONS,a.WCHTOTALCUBES,a.WCHPOSTEDTIMESTAMP,a.WCHTOTALWEIGHT,
 (CASE 
        When a.WCHDESTINATION in ('335','CNW') then trim(a.WCHORIGIN)||'-'|| trim(a.WCHCONTAINERNUMBER)||'-'||trim(a.WCHDESTINATION)||'-'||to_char(a.WCHPOSTEDTIMESTAMP,'yyyy-mm-dd') 
        ELSE trim(a.WCHORIGIN)||'-'|| trim(a.WCHCONTAINERNUMBER)||'-'||trim(a.WCHDESTINATION) END) as Container#
FROM  WWUSAF.WVCNTHD a
WHERE a.WCHCONTAINERSTATUS in ('P','T') AND a.WCHORIGIN IN ('31')  AND a.WCHPOSTEDTIMESTAMP BETWEEN char(current date - 14 days) and char(current DATE) 
and substr(trim(a.WCHCONTAINERNUMBER),1,4) NOT IN ('AAAR','AIIR','AAIR','AIRR','AIR_','AIR1')

union all

SELECT  a.WCHDOORNUMBER,a.WCHCONTAINERNUMBER,a.WCHORIGIN,a.WCHDESTINATION,a.WCHCONTAINERSTATUS,a.WCHTOTALCARTONS,a.WCHTOTALCUBES,a.WCHPOSTEDTIMESTAMP,a.WCHTOTALWEIGHT,
(CASE 
        When a.WCHDESTINATION in ('335','CNW') then trim(a.WCHORIGIN)||'-'|| trim(a.WCHCONTAINERNUMBER)||'-'||trim(a.WCHDESTINATION)||'-'||to_char(a.WCHPOSTEDTIMESTAMP,'yyyy-mm-dd') 
        ELSE trim(a.WCHORIGIN)||'-'|| trim(a.WCHCONTAINERNUMBER)||'-'||trim(a.WCHDESTINATION) END) as Container#
FROM  ASHLEYARCW.WVCNTHDA a
WHERE a.WCHCONTAINERSTATUS in ('P','T') AND a.WCHPOSTEDTIMESTAMP BETWEEN char(current date - 14 days) and char(current DATE)  and a.WCHORIGIN in ('31') and 
substr(trim(a.WCHCONTAINERNUMBER),1,4) NOT IN ('AAAR','AIIR','AAIR','AIRR','AIR_','AIR1')) as x1) as Y2


ON Y1.ContainerNumber = Y2.DistinctContainer

