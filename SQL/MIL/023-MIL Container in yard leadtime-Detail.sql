-- by product to accumulate the product as string, made by JimShen on Jul.18.2022
select a8.WCIORIGIN,a8.CONTAINER#,a8.CONTAINERTYPE,replace(replace(xml2clob(xmlagg(xmlelement(NAME A, a8.PRODUCT||','))),'<A>',''),'</A>','') as PRODUCT 
FROM
(
SELECT a9.WCIORIGIN,a9.CONTAINER#,a9.PRODUCT,a9.CONTAINERTYPE
FROM 
(-- MIL Container Utilization Details report updated on Nov.23 by Jimshen
SELECT WCIORIGIN,CONTAINER#,CUBES,ITCLS,PRODUCT,WCIDESTINATION,WCIORDER,ITEMNUMBER,QTY,WCILASTMAINTENANCETIMESTAMP,CONTAINERNUMBER,DISTINCTCONTAINER,
WCILASTMAINTENANCEUSER,ITMCQTY,UNITCUBE,UNITWEIGHT,CARTONS,CONTAINERTYPE,Date,WCHCONTAINERSIZE,
(case 
	when trim(substr(WCHCONTAINERSIZE,1,3)) = '40H' then CUBES/2650
	when trim(substr(WCHCONTAINERSIZE,1,3)) = '40' then CUBES/2383
	when trim(substr(WCHCONTAINERSIZE,1,3)) = '45' then CUBES/3058
	when substr(WCHCONTAINERSIZE,1,1) = '2' then CUBES/1191
	ELSE CUBES/2650 END) AS Utilization

FROM
(SELECT s1.ContainerNumber,s1.WCIORIGIN,s1.WCIDESTINATION,s1.WCIORDER,s1.ItemNumber,s1.Qty,s1.WCILASTMAINTENANCETIMESTAMP,s1.WCILASTMAINTENANCEUSER,s1.ITMCQTY,
s1.itcls,s1.UnitCube,s1.UnitWeight,s1.Cubes,s1.Cartons,s1.Product,s2.ContainerType,to_char(s1.WCILASTMAINTENANCETIMESTAMP,'yyyy-mm-dd') as Date, 
s2.Container#
FROM 
((SELECT trim(a.WCICONTAINERNUMBER) as ContainerNumber, a.WCIORIGIN, a.WCIDESTINATION, a.WCIORDER, trim(a.WCIITEMNUMBER) as ItemNumber, a.WCIQUANTITYLOADED as Qty, 
a.WCILASTMAINTENANCETIMESTAMP, a.WCILASTMAINTENANCEUSER, b.ITMCQTY, c.itcls,c.B2Z95S as UnitCube, c.WEGHT as UnitWeight, a.WCIQUANTITYLOADED*c.B2Z95S as Cubes,
CEIL(a.WCIQUANTITYLOADED/b.ITMCQTY) as Cartons,
trim(a.WCIORIGIN)||'-'|| trim(a.WCICONTAINERNUMBER)||'-'||trim(a.WCIDESTINATION) as Container#,
(CASE 
        WHEN c.ITCLS IN ('SLDK') THEN 'RP'
        WHEN c.ITCLS LIKE 'T%' THEN 'RP'
		WHEN c.ITCLS LIKE 'R%' THEN 'RP'
		WHEN c.ITCLS IN ('PACS') THEN 'UnKits'
        WHEN c.ITCLS LIKE 'Z%' AND c.ITCLS LIKE '%K' THEN 'UnKits'
        WHEN c.ITCLS IN ('ZACM','ZASU','ZMLH','ZMLR','ZUSR','ZUSU','ZVUC','ZXUC','ZUSU') THEN 'UPH'
        WHEN c.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB','ZECD') THEN 'CG'
        WHEN c.ITCLS IN ('ZKIS') THEN 'Bedding'		
        WHEN c.ITCLS IN ('WPLS') THEN 'Plastics'
        WHEN c.ITCLS IN ('WVBC','WVCS') THEN 'Foundation'		
		WHEN c.ITCLS IN ('PANL') THEN 'Panel'
		WHEN c.ITCLS IN ('ZKIZ') THEN 'ZipperCover'
        WHEN c.ITCLS IN ('BBFR','WVHC') THEN 'Verona'		
		WHEN c.ITCLS NOT LIKE 'Z%' THEN 'RawMaterial' 
        ELSE 'Check' END) AS Product
FROM LLUSAF.WVCNTID as a, AFILELIBL.ITMEXT as b, AMFLIBL.ITMRVA as c 
WHERE (a.WCIITEMNUMBER = b.itnbr) and a.WCIITEMNUMBER = c.itnbr and a.WCIORIGIN = c.STID and a.WCIORIGIN in('51')  
and  a.WCILASTMAINTENANCETIMESTAMP  between char(current date - 40 days) and char(current DATE)  
AND substr(trim(a.WCICONTAINERNUMBER),1,4) NOT IN ('AAAR','AIIR','AAIR','AIRR','AIR_','AIR1','AAII','ARRR')
Order by a.WCICONTAINERNUMBER, a.WCIORIGIN, a.WCILASTMAINTENANCETIMESTAMP, a.WCIDESTINATION)

union all

(SELECT trim(a.WCICONTAINERNUMBER) as ContainerNumber, a.WCIORIGIN, a.WCIDESTINATION, a.WCIORDER, trim(a.WCIITEMNUMBER) as ItemNumber, a.WCIQUANTITYLOADED as Qty, 
a.WCILASTMAINTENANCETIMESTAMP, a.WCILASTMAINTENANCEUSER, b.ITMCQTY, c.itcls,c.B2Z95S as UnitCube, c.WEGHT as UnitWeight, a.WCIQUANTITYLOADED*c.B2Z95S as Cubes,
CEIL(a.WCIQUANTITYLOADED/b.ITMCQTY) as Cartons,
trim(a.WCIORIGIN)||'-'|| trim(a.WCICONTAINERNUMBER)||'-'||trim(a.WCIDESTINATION)||'-'||SUBSTR(char(a.WCIARCHIVETIMESTAMP),1,13) as Container#,
(CASE 
        WHEN c.ITCLS IN ('SLDK') THEN 'RP'
        WHEN c.ITCLS LIKE 'T%' THEN 'RP'
		WHEN c.ITCLS LIKE 'R%' THEN 'RP'
		WHEN c.ITCLS IN ('PACS') THEN 'UnKits'
        WHEN c.ITCLS LIKE 'Z%' AND c.ITCLS LIKE '%K' THEN 'UnKits'
        WHEN c.ITCLS IN ('ZACM','ZASU','ZMLH','ZMLR','ZUSR','ZUSU','ZVUC','ZXUC','ZUSU') THEN 'UPH'
        WHEN c.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB','ZECD') THEN 'CG'
        WHEN c.ITCLS IN ('ZKIS') THEN 'Bedding'		
        WHEN c.ITCLS IN ('WPLS') THEN 'Plastics'
        WHEN c.ITCLS IN ('WVBC','WVCS') THEN 'Foundation'		
		WHEN c.ITCLS IN ('PANL') THEN 'Panel'
		WHEN c.ITCLS IN ('ZKIZ') THEN 'ZipperCover'
        WHEN c.ITCLS IN ('BBFR','WVHC') THEN 'Verona'		
		WHEN c.ITCLS NOT LIKE 'Z%' THEN 'RawMaterial' 
        ELSE 'Check' END) AS Product
FROM ASHLEYARCL.WVCNTIDA as a, AFILELIBL.ITMEXT as b, AMFLIBL.ITMRVA as c
WHERE (a.WCIITEMNUMBER = b.itnbr) and a.WCIITEMNUMBER = c.itnbr and a.WCIORIGIN = c.STID and a.WCIORIGIN in('51') and a.WCILASTMAINTENANCETIMESTAMP  
between char(current date - 40 days) and char(current DATE)  AND substr(a.WCICONTAINERNUMBER,1,4) NOT IN ('AAAR','AIIR','AAIR','AIRR','AIR_','AIR1','AAII','ARRR')
Order by a.WCICONTAINERNUMBER, a.WCIORIGIN, a.WCILASTMAINTENANCETIMESTAMP, a.WCIDESTINATION)) as s1,

-- TABLE.S2 to judge container type (combined or non-conbined)
(SELECT Container#,(case when count(Distinct PRODUCT)=1 then 'None-Mixed'  else 'Mixed' end) as ContainerType
FROM 
((SELECT trim(a.WCICONTAINERNUMBER) as ContainerNumber, a.WCIORIGIN, a.WCIDESTINATION, a.WCIORDER, trim(a.WCIITEMNUMBER) as ItemNumber, a.WCIQUANTITYLOADED as Qty, 
a.WCILASTMAINTENANCETIMESTAMP, a.WCILASTMAINTENANCEUSER, b.ITMCQTY, c.itcls,c.B2Z95S as UnitCube, c.WEGHT as UnitWeight, a.WCIQUANTITYLOADED*c.B2Z95S as Cubes,
CEIL(a.WCIQUANTITYLOADED/b.ITMCQTY) as Cartons,
trim(a.WCIORIGIN)||'-'|| trim(a.WCICONTAINERNUMBER)||'-'||trim(a.WCIDESTINATION) as Container#,
(CASE 
        WHEN c.ITCLS IN ('SLDK') THEN 'RP'
        WHEN c.ITCLS LIKE 'T%' THEN 'RP'
		WHEN c.ITCLS LIKE 'R%' THEN 'RP'
		WHEN c.ITCLS IN ('PACS') THEN 'UnKits'
        WHEN c.ITCLS LIKE 'Z%' AND c.ITCLS LIKE '%K' THEN 'UnKits'
        WHEN c.ITCLS IN ('ZACM','ZASU','ZMLH','ZMLR','ZUSR','ZUSU','ZVUC','ZXUC','ZUSU') THEN 'UPH'
        WHEN c.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB','ZECD') THEN 'CG'
        WHEN c.ITCLS IN ('ZKIS') THEN 'Bedding'		
        WHEN c.ITCLS IN ('WPLS') THEN 'Plastics'
        WHEN c.ITCLS IN ('WVBC','WVCS') THEN 'Foundation'		
		WHEN c.ITCLS IN ('PANL') THEN 'Panel'
		WHEN c.ITCLS IN ('ZKIZ') THEN 'ZipperCover'
        WHEN c.ITCLS IN ('BBFR','WVHC') THEN 'Verona'		
		WHEN c.ITCLS NOT LIKE 'Z%' THEN 'RawMaterial' 
        ELSE 'Check' END) AS Product
FROM LLUSAF.WVCNTID as a, AFILELIBL.ITMEXT as b, AMFLIBL.ITMRVA as c
WHERE (a.WCIITEMNUMBER = b.itnbr) and a.WCIITEMNUMBER = c.itnbr and a.WCIORIGIN = c.STID and a.WCIORIGIN in('51')  
and  a.WCILASTMAINTENANCETIMESTAMP BETWEEN char(current date - 40 days) and char(current DATE) 
AND substr(trim(a.WCICONTAINERNUMBER),1,4) NOT IN ('AAAR','AIIR','AAIR','AIRR','AIR_','AIR1','AAII','ARRR')
Order by a.WCICONTAINERNUMBER, a.WCIORIGIN, a.WCILASTMAINTENANCETIMESTAMP, a.WCIDESTINATION)

union all

(SELECT trim(a.WCICONTAINERNUMBER) as ContainerNumber, a.WCIORIGIN, a.WCIDESTINATION, a.WCIORDER, trim(a.WCIITEMNUMBER) as ItemNumber, a.WCIQUANTITYLOADED as Qty, 
a.WCILASTMAINTENANCETIMESTAMP, a.WCILASTMAINTENANCEUSER, b.ITMCQTY, c.itcls,c.B2Z95S as UnitCube, c.WEGHT as UnitWeight, a.WCIQUANTITYLOADED*c.B2Z95S as Cubes,
CEIL(a.WCIQUANTITYLOADED/b.ITMCQTY) as Cartons,
trim(a.WCIORIGIN)||'-'|| trim(a.WCICONTAINERNUMBER)||'-'||trim(a.WCIDESTINATION)||'-'||SUBSTR(char(a.WCIARCHIVETIMESTAMP),1,13) as Container#,
(CASE 
        WHEN c.ITCLS IN ('SLDK') THEN 'RP'
        WHEN c.ITCLS LIKE 'T%' THEN 'RP'
		WHEN c.ITCLS LIKE 'R%' THEN 'RP'
		WHEN c.ITCLS IN ('PACS') THEN 'UnKits'
        WHEN c.ITCLS LIKE 'Z%' AND c.ITCLS LIKE '%K' THEN 'UnKits'
        WHEN c.ITCLS IN ('ZACM','ZASU','ZMLH','ZMLR','ZUSR','ZUSU','ZVUC','ZXUC','ZUSU') THEN 'UPH'
        WHEN c.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB','ZECD') THEN 'CG'
        WHEN c.ITCLS IN ('ZKIS') THEN 'Bedding'		
        WHEN c.ITCLS IN ('WPLS') THEN 'Plastics'
        WHEN c.ITCLS IN ('WVBC','WVCS') THEN 'Foundation'		
		WHEN c.ITCLS IN ('PANL') THEN 'Panel'
		WHEN c.ITCLS IN ('ZKIZ') THEN 'ZipperCover'
        WHEN c.ITCLS IN ('BBFR','WVHC') THEN 'Verona'		
		WHEN c.ITCLS NOT LIKE 'Z%' THEN 'RawMaterial' 
        ELSE 'Check' END) AS Product
FROM ASHLEYARCL.WVCNTIDA as a, AFILELIBL.ITMEXT as b, AMFLIBL.ITMRVA as c
WHERE (a.WCIITEMNUMBER = b.itnbr) and a.WCIITEMNUMBER = c.itnbr and a.WCIORIGIN = c.STID and a.WCIORIGIN in('51') and a.WCILASTMAINTENANCETIMESTAMP 
between char(current date - 40 days) and char(current DATE) AND substr(a.WCICONTAINERNUMBER,1,4) NOT IN ('AAAR','AIIR','AAIR','AIRR','AIR_','AIR1','AAII','ARRR')
Order by a.WCICONTAINERNUMBER, a.WCIORIGIN, a.WCILASTMAINTENANCETIMESTAMP, a.WCIDESTINATION))
group by WCIORIGIN, Container#
order by Container#) as s2

WHERE s1.Container# = s2.Container#
order by s1.WCIORIGIN,s1.ContainerNumber,s1.WCILASTMAINTENANCETIMESTAMP
) AS Y1
-- TABLE Y1 to get current and archived container loaded details

right join

(Select DISTINCT(x1.Container#) as DistinctContainer,WCHCONTAINERSIZE

FROM
(SELECT 
a.WCHCONTAINERNUMBER,a.WCHORIGIN,a.WCHDESTINATION,a.WCHCONTAINERSTATUS,a.WCHTOTALCARTONS,a.WCHTOTALCUBES,a.WCHPOSTEDTIMESTAMP,a.WCHTOTALWEIGHT,
a.WCHCONTAINERSIZE,
trim(a.WCHORIGIN)||'-'|| trim(a.WCHCONTAINERNUMBER)||'-'||trim(a.WCHDESTINATION) as Container#
FROM  LLUSAF.WVCNTHD a
WHERE a.WCHCONTAINERSTATUS in ('A','P','T','H','R') AND a.WCHORIGIN IN ('51')  AND a.WCHPOSTEDTIMESTAMP BETWEEN char(current date - 30 days) and char(current DATE) 
and substr(trim(a.WCHCONTAINERNUMBER),1,4) NOT IN ('AAAR','AIIR','AAIR','AIRR','AIR_','AIR1','AAII','ARRR')

union all

SELECT  a.WCHCONTAINERNUMBER,a.WCHORIGIN,a.WCHDESTINATION,a.WCHCONTAINERSTATUS,a.WCHTOTALCARTONS,a.WCHTOTALCUBES,a.WCHPOSTEDTIMESTAMP,a.WCHTOTALWEIGHT,a.WCHCONTAINERSIZE,
trim(a.WCHORIGIN)||'-'|| trim(a.WCHCONTAINERNUMBER)||'-'||trim(a.WCHDESTINATION)||'-'||SUBSTR(char(a.WCHARCHIVETIMESTAMP),1,13) as Container#
FROM  ASHLEYARCL.WVCNTHDA a
WHERE a.WCHCONTAINERSTATUS in ('P','T') AND a.WCHPOSTEDTIMESTAMP BETWEEN char(current date - 30 days) and char(current DATE)  and a.WCHORIGIN in ('51') and 
substr(trim(a.WCHCONTAINERNUMBER),1,4) NOT IN ('AAAR','AIIR','AAIR','AIRR','AIR_','AIR1','AAII','ARRR')) as x1) as Y2
ON Y1.Container# = Y2.DistinctContainer
) a9
GROUP BY a9.WCIORIGIN,a9.CONTAINER#,a9.PRODUCT,a9.CONTAINERTYPE
) a8
GROUP BY a8.WCIORIGIN,a8.CONTAINER#,a8.CONTAINERTYPE