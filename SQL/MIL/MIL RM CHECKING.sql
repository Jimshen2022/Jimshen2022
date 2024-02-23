
-- MIL UnKits Output Tracking. created on Aug.30.2021

-- OPITON 1365 RM UNKITS LEFT JOIN SCANNED INTO CONTAINER'S SN STATUS
SELECT X1.AACOD1,X1.AATWHS,X1.AAITM#,X1.AAADAT, X1.AAATIM,X1.AAAUSR,X1.AAAPGM,X1.AAORD#,X1.AAVND#,X1.AAEMP#,X1.SN,X2.CNT_STATUS,
X1.SHIFT,right('000000'||ltrim(X1.AAATIM),6) AS TIME,
X2.WCSCONTAINERNUMBER,
(CASE WHEN X2.CNT_STATUS IS NULL THEN 'InWhse Or PackingLine' ELSE X2.CNT_STATUS END) AS STATUS

FROM
(
-- OPTION 1365 TO GET UN-KITS OUTPUT 
SELECT AACOD1,AATWHS,AAITM#,AAADAT, AAATIM,AAAUSR,AAAPGM,AAORD#,AAVND#,AAEMP#, char(AASER#) as SN,
(CASE WHEN t1.AAATIM BETWEEN '070000' AND '194459' THEN 'DS' ELSE 'NS' END) AS SHIFT
FROM DISTLIBL.ACTAUDT t1
WHERE (t1.AAADAT 
BETWEEN int(substr(trim(char(CURRENT DATE - 45 days)),1,4)||substr(trim(char(CURRENT DATE- 45 days)),6,2)||substr(trim(char(CURRENT DATE- 45 days)),9,2)) 
AND int(substr(trim(char(CURRENT DATE + 1 days)),1,4)||substr(trim(char(CURRENT DATE + 1 days)),6,2)||substr(trim(char(CURRENT DATE + 1 days)),9,2))) 

AND (t1.AATWHS='51') and (t1.AACOD1 = 'RM') AND (AASER# >0 ) AND TRIM(t1.AAITM#) LIKE '%UN' AND TRIM(t1.AAITM#) NOT LIKE 'M%'
) AS X1 


LEFT JOIN

-- CURRENT AND ARCHIVED UnKits SN SA done
(
(SELECT a.WCSSERIALNUMBER,a.WCSCONTAINERNUMBER, b.WCHCONTAINERSTATUS,'SADone' as CNT_STATUS
FROM LLUSAF.WVCNTSD a, LLUSAF.WVCNTHD b 
WHERE trim(a.WCSITEMNUMBER) LIKE '%UN' AND trim(a.WCSITEMNUMBER) NOT LIKE 'M%' and a.WCSCONTAINERNUMBER = b.WCHCONTAINERNUMBER and b.WCHCONTAINERSTATUS IN ('P','T') AND b.WCHPOSTEDTIMESTAMP BETWEEN CHAR(CURRENT DATE - 60 days) AND CHAR(CURRENT DATE + 1 days))
UNION
-- Archived container serial numbers in past 60 days
(SELECT a.WCSSERIALNUMBER, a.WCSCONTAINERNUMBER,'P&T' as WCHCONTAINERSTATUS,'SADone' as CNT_STATUS
FROM ASHLEYARCL.WVCNTSDA a
WHERE TRIM(a.WCSITEMNUMBER) LIKE '%UN' AND TRIM(a.WCSITEMNUMBER) NOT LIKE 'M%' AND a.WCSADDEDTIMESTAMP between char(current date - 120 days) and char(current DATE + 1 days))
UNION ALL
-- CURRNT NOT SA SN STATUS
SELECT a.WCSSERIALNUMBER,a.WCSCONTAINERNUMBER,b.WCHCONTAINERSTATUS,
(CASE 
    WHEN a.WCSCONTAINERNUMBER LIKE 'MRUN%' THEN 'InTempCTN'
	WHEN a.WCSCONTAINERNUMBER LIKE 'KECR%' THEN 'InTempCTN'
	WHEN a.WCSCONTAINERNUMBER LIKE 'KHOA%' THEN 'InTempCTN'
	WHEN a.WCSCONTAINERNUMBER LIKE 'M3K%' THEN 'InTempCTN'
	WHEN a.WCSCONTAINERNUMBER LIKE 'RUN%' THEN 'InTempCTN'
	WHEN b.WCHCONTAINERSTATUS IN ('A','C','H','R','U') THEN 'InRealCTN'
	ELSE 'CheckCNTStatus' END) AS CNT_STATUS
	
FROM LLUSAF.WVCNTSD a, LLUSAF.WVCNTHD b 
WHERE trim(a.WCSITEMNUMBER) LIKE '%UN' AND trim(a.WCSITEMNUMBER) NOT LIKE 'M%' and a.WCSCONTAINERNUMBER = b.WCHCONTAINERNUMBER and b.WCHCONTAINERSTATUS NOT IN ('P','T') AND b.WCHPOSTEDTIMESTAMP BETWEEN CHAR(CURRENT DATE - 60 days) AND CHAR(CURRENT DATE + 1 days)
) AS X2

ON X1.SN = CHAR(X2.WCSSERIALNUMBER)




/*
(
-- Archived container serial numbers in past 60 days
select a.WCSSERIALNUMBER  
From ASHLEYARCL.WVCNTSDA a
where a.WCSADDEDTIMESTAMP between char(current date - 10 days) and char(current DATE)
)



--RM

SELECT a.aaser#
FROM  DISTLIBL.ACTAUDT a
WHERE  a.AAITM# = '3540225MRUN'


--ARCHIVED SN  ASHLEYARCL.WVCNTSDA-SN
Select *
From ASHLEYARCL.WVCNTSDA a



-- CURRENT SN  LLUSAF.WVCNTSDA-SN
Select *
From LLUSAF.WVCNTSD a
limit 100



---- CURRENT CONTAINER HEAD TABLE FOR  LLUSAF.WVCNTHD-Head
Select *
From LLUSAF.WVCNTHD a
limit 100

*/
