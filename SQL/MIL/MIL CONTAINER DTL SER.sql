
-- OPITON 1365 RM UNKITS LEFT JOIN SCANNED INTO CONTAINER'S SN STATUS
SELECT X1.AACOD1,X1.AATWHS,X1.AAITM#,X1.AAADAT, X1.AAATIM,X1.AAAUSR,X1.AAAPGM,X1.AAORD#,X1.AAVND#,X1.AAEMP#,X1.SN,X2.CNT_STATUS,X2.WCSCONTAINERNUMBER,
(CASE WHEN X2.CNT_STATUS IS NULL THEN 'InWhse Or PackingLine' ELSE X2.CNT_STATUS END) AS STATUS,
(CASE WHEN X1.AAATIM BETWEEN '070000' AND '194459' THEN 'DS' ELSE 'NS' END) AS SHIFT
FROM
(
-- OPTION 1365 TO GET UN-KITS OUTPUT 
SELECT AACOD1,AATWHS,AAITM#,AAADAT, AAATIM,AAAUSR,AAAPGM,AAORD#,AAVND#,AAEMP#, char(AASER#) as SN,
(CASE WHEN t1.AAATIM BETWEEN '070000' AND '194459' THEN 'DS' ELSE 'NS' END) AS SHIFT
FROM DISTLIBL.ACTAUDT t1
WHERE (t1.AAADAT Between ? And ?) AND (t1.AATWHS='51') and (AACOD1 = 'RM') AND (AASER# >0 ) AND TRIM(t1.AAITM#) LIKE '%UN' AND TRIM(t1.AAITM#) NOT LIKE 'M%'
) AS X1 

LEFT JOIN

(
--- SCANNED INTO CONTAINER'S SN STATUS
SELECT S1.WCSCONTAINERNUMBER,S1.WCSORIGIN,S1.WCSDESTINATION,S1.WCSITEMNUMBER,CHAR(S1.SN) SN,S1.WCSADDEDTIMESTAMP,s2.WCHCONTAINERSTATUS,
(CASE 
    WHEN S1.WCSCONTAINERNUMBER LIKE 'MRUN%' THEN 'InTempCTN'
	WHEN S1.WCSCONTAINERNUMBER LIKE 'KECR%' THEN 'InTempCTN'
	WHEN S1.WCSCONTAINERNUMBER LIKE 'KHOA%' THEN 'InTempCTN'
	WHEN S1.WCSCONTAINERNUMBER LIKE 'M3K%' THEN 'InTempCTN'
	WHEN S1.WCSCONTAINERNUMBER LIKE 'RUN%' THEN 'InTempCTN'
	WHEN S2.WCHCONTAINERSTATUS IN ('P','T') THEN 'SADone'
	WHEN S2.WCHCONTAINERSTATUS IN ('A') THEN 'InRealCTN'
	ELSE 'CHECK' END) AS CNT_STATUS

FROM
-- PULL OUT SCANNED INTO CONTAINER Serial number 
(SELECT 
t1.WCSCONTAINERNUMBER,t1.WCSORIGIN,t1.WCSDESTINATION,t1.WCSITEMNUMBER,CHAR(t1.WCSSERIALNUMBER) SN,t1.WCSADDEDTIMESTAMP
FROM  LLUSAF. TBL_WVCONTAINER_DTL_SER t1
Where t1.WCSADDEDTIMESTAMP between char(current date - 30 days) and char(current DATE + 1 days) 
and trim(t1.WCSITEMNUMBER) LIKE '%UN' AND trim(t1.WCSITEMNUMBER) NOT LIKE 'M%'

UNION
-- PULL OUT ARCHIVED SCANNED INTO CONTAINER Serial number 
SELECT  
t2.WCSCONTAINERNUMBER,t2.WCSORIGIN,t2.WCSDESTINATION,t2.WCSITEMNUMBER,CHAR(t2.WCSSERIALNUMBER) SN,t2.WCSADDEDTIMESTAMP
FROM  ASHLEYARCL.TBL_WVCONTAINER_DTL_SER_A t2
Where t2.WCSADDEDTIMESTAMP between char(current date - 30 days) and char(current DATE + 1 days ) 
and trim(t2.WCSITEMNUMBER) LIKE '%UN' AND trim(t2.WCSITEMNUMBER) NOT LIKE 'M%') AS S1

LEFT JOIN
(
-- QUERY 1020.02.13 ALL CONTAINERS STATUS
SELECT 
a.WCHCONTAINERNUMBER, a.WCHCONTAINERSTATUS
FROM  LLUSAF.WVCNTHD a
) AS S2
ON  S1.WCSCONTAINERNUMBER = S2.WCHCONTAINERNUMBER
) AS X2

ON X1.SN = X2.SN






