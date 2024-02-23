
-- MIL SA DONE BUT NOT RM on OCT.21.2021 by JimShen
SELECT *
FROM 
(
(SELECT a.WCSSERIALNUMBER,a.WCSCONTAINERNUMBER, b.WCHCONTAINERSTATUS,'SADone' as CNT_STATUS
FROM LLUSAF.WVCNTSD a, LLUSAF.WVCNTHD b 
WHERE trim(a.WCSITEMNUMBER) LIKE '%UN' AND trim(a.WCSITEMNUMBER) LIKE 'M%' and a.WCSCONTAINERNUMBER = b.WCHCONTAINERNUMBER and b.WCHCONTAINERSTATUS IN ('P','T') AND b.WCHPOSTEDTIMESTAMP BETWEEN CHAR(CURRENT DATE - 60 days) AND CHAR(CURRENT DATE + 1 days))
UNION
-- Archived container serial numbers in past 60 days
(SELECT a.WCSSERIALNUMBER, a.WCSCONTAINERNUMBER,'P&T' as WCHCONTAINERSTATUS,'SADone' as CNT_STATUS
FROM ASHLEYARCL.WVCNTSDA a
WHERE TRIM(a.WCSITEMNUMBER) LIKE '%UN' AND TRIM(a.WCSITEMNUMBER) LIKE 'M%' AND a.WCSADDEDTIMESTAMP between char(current date - 120 days) and char(current DATE + 1 days))
UNION ALL
-- CURRNT NOT SA SN STATUS
SELECT a.WCSSERIALNUMBER,a.WCSCONTAINERNUMBER,b.WCHCONTAINERSTATUS,
(CASE 
    WHEN a.WCSCONTAINERNUMBER LIKE 'MRUN%' THEN 'InTempCTN'
	WHEN a.WCSCONTAINERNUMBER LIKE 'KECR%' THEN 'InTempCTN'
	WHEN a.WCSCONTAINERNUMBER LIKE 'KHO%' THEN 'InTempCTN'
	WHEN a.WCSCONTAINERNUMBER LIKE 'M3K%' THEN 'InTempCTN'
	WHEN a.WCSCONTAINERNUMBER LIKE 'M3E%' THEN 'InTempCTN'
	WHEN a.WCSCONTAINERNUMBER LIKE 'M3H%' THEN 'InTempCTN'
	WHEN a.WCSCONTAINERNUMBER LIKE 'RUN%' THEN 'InTempCTN'
	WHEN b.WCHCONTAINERSTATUS IN ('A','C','H','R','U') THEN 'InRealCTN'
	ELSE 'CheckCNTStatus' END) AS CNT_STATUS
	
FROM LLUSAF.WVCNTSD a, LLUSAF.WVCNTHD b 
WHERE trim(a.WCSITEMNUMBER) LIKE '%UN' AND trim(a.WCSITEMNUMBER) LIKE 'M%' and a.WCSCONTAINERNUMBER = b.WCHCONTAINERNUMBER and b.WCHCONTAINERSTATUS NOT IN ('P','T')
) AS X1
WHERE NOT EXISTS 
(
-- OPTION 1365 TO GET UN-KITS OUTPUT 
SELECT char(AASER#) as SN

FROM DISTLIBL.ACTAUDT t1
WHERE (t1.AAADAT 
BETWEEN int(substr(trim(char(CURRENT DATE - 145 days)),1,4)||substr(trim(char(CURRENT DATE- 145 days)),6,2)||substr(trim(char(CURRENT DATE- 145 days)),9,2)) 
AND int(substr(trim(char(CURRENT DATE + 1 days)),1,4)||substr(trim(char(CURRENT DATE + 1 days)),6,2)||substr(trim(char(CURRENT DATE + 1 days)),9,2))) 
AND (t1.AATWHS='51') and (t1.AACOD1 = 'RM') AND (AASER# >0 ) AND TRIM(t1.AAITM#) LIKE '%UN' AND TRIM(t1.AAITM#) LIKE 'M%'
)

