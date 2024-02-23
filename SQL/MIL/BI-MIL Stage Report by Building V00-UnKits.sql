-- MIL Stage report by building, created on Nov.04.2021 by JimShen
-- Pullout received but still not SA serial number
Select Y1.SN,Y1.TDTSTS,Y1.TDITEM,Y1.MO,Y1.TDWHSE,Y1.TDMDAT,Y1.TDMTME,Y1.TXT_TIME,Y1.ITCLS,(1/Y1.ITMCQTY) as Cartons, 
Y1.Product,Y1.SHIFT,Y1.Line,Y2.CTN,Y2."CTN_Status"
From
(
SELECT X1.SN,X1.TDTSTS,X1.TDITEM,X1.MO,X1.TDWHSE,X1.TDMDAT,X1.TDMTME,X1.TXT_TIME,X1.ITCLS,X1.Product,X1.SHIFT,CHAR(TRIM(SUBSTR(X2.JOBNO,1,5))) AS Line,X1.ITMCQTY
FROM
( 
SELECT CHAR(trim(t1.TDTAG#)) AS SN,t1.TDITEM,t1.TDAPO# as MO,t1.TDWHSE,t1.TDMDAT,t1.TDMTME,
t1.TDTSTS,right('000000'||ltrim(t1.TDMTME),6) AS TXT_TIME,t3.ITCLS,t5.ITMCQTY,
(CASE 
        WHEN t3.ITCLS IN ('WPLS') THEN 'Plastics'
        WHEN t3.ITCLS IN ('WVBC','WVHC') THEN 'Foundation'
        WHEN t3.ITCLS IN ('SLDK') THEN 'RP'
        WHEN t3.ITCLS LIKE 'T%' THEN 'RP'
        WHEN t3.ITCLS IN ('ZKIS') THEN 'Bedding'
        WHEN t3.ITCLS IN ('ZKIZ') THEN 'ZipperCover'
        WHEN t3.ITCLS LIKE 'Z%' AND t3.ITCLS LIKE '%K' THEN 'UnKits'
        WHEN t3.ITCLS IN ('PACS') THEN 'UnKits'
        WHEN t3.ITCLS IN ('BBFR') THEN 'FR SOCK'
        WHEN t3.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB') THEN 'CG'
        WHEN t3.ITCLS LIKE 'Z%' THEN 'UPH'
        ELSE 'Check' END) AS Product,
(CASE WHEN t1.TDMTME BETWEEN '070000' AND '194459' THEN 'DS' ELSE 'NS' END) AS SHIFT

FROM DISTLIBL.TAGINVD t1,(SELECT DISTINCT t2.ITNBR,t2.ITCLS FROM AMFLIBL.ITEMBL t2 WHERE t2.HOUSE = '51' GROUP BY t2.ITNBR,t2.ITCLS) AS t3,
(SELECT DISTINCT t4.ITNBR,t4.ITMCQTY FROM AFILELIBL.ITMEXT t4 GROUP BY t4.ITNBR,t4.ITMCQTY) AS t5

WHERE t1.TDITEM = t3.ITNBR AND t1.TDITEM=t5.ITNBR AND t3.ITNBR=t5.ITNBR and t1.TDTSTS IN ('R','S') and t1.TDMDAT 
BETWEEN  int('1'||substr(trim(char(CURRENT DATE - 90 days)),3,2)||substr(trim(char(CURRENT DATE- 90 days)),6,2)||substr(trim(char(CURRENT DATE- 90 days)),9,2)) 
AND int('1'||substr(trim(char(CURRENT DATE + 1 days)),3,2)||substr(trim(char(CURRENT DATE + 1 days)),6,2)||substr(trim(char(CURRENT DATE + 1 days)),9,2))
AND (t3.ITCLS LIKE 'Z%K' OR t3.ITCLS IN ('PACS'))	
AND CHAR(TRIM(t1.TDTAG#)) NOT IN
		-- CURRENT AND ARCHIVED F/G SN SA done
		(SELECT CHAR(trim(a.WCSSERIALNUMBER)) AS SN
			FROM LLUSAF.WVCNTSD a, LLUSAF.WVCNTHD b 
			WHERE a.WCSCONTAINERNUMBER = b.WCHCONTAINERNUMBER and b.WCHCONTAINERSTATUS IN ('P','T') AND b.WCHPOSTEDTIMESTAMP BETWEEN CHAR(CURRENT DATE - 91 days) AND CHAR(CURRENT DATE + 1 days))

AND CHAR(TRIM(t1.TDTAG#)) NOT IN 
			-- Archived container serial numbers in past 90 days
		(SELECT CHAR(TRIM(a.WCSSERIALNUMBER)) AS SN
			FROM ASHLEYARCL.WVCNTSDA a
			WHERE a.WCSADDEDTIMESTAMP between char(current date - 91 days) and char(current DATE + 1 days)) 

) AS X1

LEFT JOIN
(
-- pull out the MO and Line information
(SELECT t1.ORDNO,t1.JOBNO,t1.FITEM FROM AMFLIBL.MOMAST t1 where t1.CRDT
BETWEEN  int('1'||substr(trim(char(CURRENT DATE - 180 days)),3,2)||substr(trim(char(CURRENT DATE- 180 days)),6,2)||substr(trim(char(CURRENT DATE- 180 days)),9,2)) 
AND int('1'||substr(trim(char(CURRENT DATE + 1 days)),3,2)||substr(trim(char(CURRENT DATE + 1 days)),6,2)||substr(trim(char(CURRENT DATE + 1 days)),9,2))
order by t1.CRDT DESC
) 

UNION ALL

(SELECT t1.ORDNO,t1.JOBNO,t1.FITEM FROM AMFLIBL.MOHMST t1 where t1.CRDT
BETWEEN  int('1'||substr(trim(char(CURRENT DATE - 180 days)),3,2)||substr(trim(char(CURRENT DATE- 180 days)),6,2)||substr(trim(char(CURRENT DATE- 180 days)),9,2)) 
AND int('1'||substr(trim(char(CURRENT DATE + 1 days)),3,2)||substr(trim(char(CURRENT DATE + 1 days)),6,2)||substr(trim(char(CURRENT DATE + 1 days)),9,2))
order by t1.CRDT DESC
)
) AS X2
ON X1.MO=X2.ORDNO and X1.TDITEM=X2.FITEM
) as Y1

Left join
-- Pull out SN scanned into temp ctn and real ctn
(
SELECT CHAR(trim(a.WCSSERIALNUMBER)) AS SN,a.WCSCONTAINERNUMBER AS CTN,
	(CASE 
		WHEN a.WCSCONTAINERNUMBER LIKE 'MRUN%' THEN 'InTempCTN'
		WHEN a.WCSCONTAINERNUMBER LIKE 'KECR%' THEN 'InTempCTN'
		WHEN a.WCSCONTAINERNUMBER LIKE 'KHO%' THEN 'InTempCTN'
		WHEN a.WCSCONTAINERNUMBER LIKE 'M3K%' THEN 'InTempCTN'
		WHEN a.WCSCONTAINERNUMBER LIKE 'M3E%' THEN 'InTempCTN'
		WHEN a.WCSCONTAINERNUMBER LIKE 'M3H%' THEN 'InTempCTN'
		WHEN a.WCSCONTAINERNUMBER LIKE 'RUN%' THEN 'InTempCTN'
		ELSE 'InRealCTN' END) AS "CTN_Status"
						
FROM LLUSAF.WVCNTSD a, LLUSAF.WVCNTHD b 
WHERE a.WCSCONTAINERNUMBER = b.WCHCONTAINERNUMBER and b.WCHCONTAINERSTATUS NOT IN ('P','T') 

) AS Y2
ON Y1.SN = Y2.SN
