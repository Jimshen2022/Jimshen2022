
-- MIL Container Loading Scan (SA finished and un_SA containers), created by Jimshen on Sep.27.2021

select t1.WCICNTNR, t1.WCIORIGN, t1.WCIDESTN, t1.WCIORDRN, t1.WCIITNBR, t1.WCIQTLOD, t1.WCILSMTS, t1.WCILSMUS, ROUND(t1.WCIQTLOD/t2.ITMCQTY) as BoxQty, t1.WCIQTLOD*t3.B2Z95S as CUBES, t3.ITCLS,
(CASE 
        WHEN t3.ITCLS IN ('WPLS') THEN 'Plastics'
        WHEN t3.ITCLS IN ('PANL') THEN 'Panel'
        WHEN t3.ITCLS IN ('DECK','QA') THEN 'RawMaterial'
        WHEN t3.ITCLS IN ('WVBC','WVHC','WVCS') THEN 'Foundation'
        WHEN t3.ITCLS IN ('SLDK') THEN 'RP'
        WHEN t3.ITCLS LIKE 'T%' THEN 'RP'
        WHEN t3.ITCLS IN ('ZKIS') THEN 'Bedding'
        WHEN t3.ITCLS IN ('ZKIZ') THEN 'ZipperCover'
        WHEN t3.ITCLS LIKE 'Z%' AND t3.ITCLS LIKE '%K' THEN 'UnKits'
        WHEN t3.ITCLS IN ('PACS') THEN 'UnKits'
        WHEN t3.ITCLS IN ('BBFR') THEN 'Verona'
        WHEN t3.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB') THEN 'CG'
        WHEN t3.ITCLS LIKE 'Z%' THEN 'UPH'
        ELSE 'Check' END) AS Product,
(CASE WHEN char(substr(char(TIME(t1.WCILSMTS)),1,2)||substr(char(TIME(t1.WCILSMTS)),4,2)||
substr(char(TIME(t1.WCILSMTS)),7,2)) BETWEEN '070000' AND '194459' THEN 'DS' ELSE 'NS' END) AS SHIFT

From LLUSAF.WVCNTID t1, AFILELIBL.ITMEXT t2, AMFLIBL.ITMRVA t3
where t1.WCILSMTS 
between timestamp(trim(char(CURRENT DATE - 1 days))||'-07.00.00.000000') and  timestamp(trim(char(CURRENT DATE))||'-06.59.59.999999')
and t1.WCIITNBR= t2.ITNBR 
and t1.wcicntnr not LIKE 'KECR%'
and t1.wcicntnr not LIKE 'M3K%'
and t1.wcicntnr not LIKE 'M3E%'
and t1.wcicntnr not LIKE 'M3H%'
and t1.wcicntnr not LIKE 'KHO%'
and t1.wcicntnr not LIKE 'MRUN%'
and t1.wcicntnr not LIKE 'RUN%'
and t1.WCIITNBR = t3.itnbr and t1.wciorign=t3.stid
order by  t1.WCICNTNR ,t1.WCILSMTS




-- MIL Container Loading Scan (SA finished and un_SA containers), created by Jimshen on Sep.27.2021

select t1.WCICNTNR, t1.WCIORIGN, t1.WCIDESTN, t1.WCIORDRN, t1.WCIITNBR, t1.WCIQTLOD, t1.WCILSMTS, t1.WCILSMUS, ROUND(t1.WCIQTLOD/t2.ITMCQTY) as BoxQty, t1.WCIQTLOD*t3.B2Z95S as CUBES, t3.ITCLS,
(CASE 
        WHEN t3.ITCLS IN ('WPLS') THEN 'Plastics'
        WHEN t3.ITCLS IN ('PANL') THEN 'Panel'
        WHEN t3.ITCLS IN ('DECK','QA') THEN 'RawMaterial'
        WHEN t3.ITCLS IN ('WVBC','WVHC','WVCS') THEN 'Foundation'
        WHEN t3.ITCLS IN ('SLDK') THEN 'RP'
        WHEN t3.ITCLS LIKE 'T%' THEN 'RP'
        WHEN t3.ITCLS IN ('ZKIS') THEN 'Bedding'
        WHEN t3.ITCLS IN ('ZKIZ') THEN 'ZipperCover'
        WHEN t3.ITCLS LIKE 'Z%' AND t3.ITCLS LIKE '%K' THEN 'UnKits'
        WHEN t3.ITCLS IN ('PACS') THEN 'UnKits'
        WHEN t3.ITCLS IN ('BBFR') THEN 'Verona'
        WHEN t3.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB') THEN 'CG'
        WHEN t3.ITCLS LIKE 'Z%' THEN 'UPH'
        ELSE 'Check' END) AS Product,
(CASE WHEN char(substr(char(TIME(t1.WCILSMTS)),1,2)||substr(char(TIME(t1.WCILSMTS)),4,2)||
substr(char(TIME(t1.WCILSMTS)),7,2)) BETWEEN '070000' AND '194459' THEN 'DS' ELSE 'NS' END) AS SHIFT

From LLUSAF.WVCNTID t1, AFILELIBL.ITMEXT t2, AMFLIBL.ITMRVA t3
where t1.WCILSMTS 
between ? and  ?
and t1.WCIITNBR= t2.ITNBR 
and t1.wcicntnr not LIKE 'KECR%'
and t1.wcicntnr not LIKE 'M3K%'
and t1.wcicntnr not LIKE 'M3E%'
and t1.wcicntnr not LIKE 'M3H%'
and t1.wcicntnr not LIKE 'KHO%'
and t1.wcicntnr not LIKE 'MRUN%'
and t1.wcicntnr not LIKE 'RUN%'
and t1.WCIITNBR = t3.itnbr and t1.wciorign=t3.stid
order by  t1.WCICNTNR ,t1.WCILSMTS