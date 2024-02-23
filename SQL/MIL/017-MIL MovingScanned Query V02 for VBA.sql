-- MIL Moving Scanned SN By Shift create by Jimshen on Jun.03.2022

SELECT X1.SN, 1 as QTY,X1.AACOD1,X1.AATWHS,X1.Location,Y1.AAORD#,X1.AAITM#,X1.AAEMP#,X1.AAAUSR, X1.SHIFT, X1.ITMCQTY, X1.ScannedTime,HOUR(X1.ScannedTime) AS HOUR,X1.Cartons,X1.ITCLS,X1.Product
FROM (SELECT  CHAR(a.AASER#) as SN,a.AACOD1,a.AATWHS,a.AATARA||'00'||a.AATASL||a.AATSEC||a.AATTIR as Location, a.AAORD#,a.AAITM#, a.AAEMP#,a.AAAUSR,
(CASE WHEN char(substr(char(TIME(to_date(a.AAADAT||' '||right('000000'||ltrim(a.AAATIM),6), 'yyyymmdd hh24:mi:ss'))),1,2)||substr(char(TIME(to_date(a.AAADAT||' '||right('000000'||ltrim(a.AAATIM),6), 'yyyymmdd hh24:mi:ss'))),4,2)||substr(char(TIME(to_date(a.AAADAT||' '||right('000000'||ltrim(a.AAATIM),6), 'yyyymmdd hh24:mi:ss'))),7,2)) BETWEEN '070000' AND '194459' THEN 'DS' ELSE 'NS' END) AS SHIFT,
t2.ITMCQTY,t3.ITCLS,
(CASE WHEN t3.ITCLS LIKE 'TAF%' THEN 'RP' WHEN t3.ITCLS IN ('PACS') THEN 'UnKits' WHEN t3.ITCLS LIKE 'Z%' AND t3.ITCLS LIKE '%K' THEN 'UnKits' WHEN t3.ITCLS IN ('ZACM','ZASU','ZMLH','ZMLR','ZUSR','ZUSU','ZVUC','ZXUC','ZUSU','ZUMU') THEN 'UPH' WHEN t3.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB') THEN 'CG'  WHEN t3.ITCLS IN ('ZKIS') THEN 'Bedding'	 WHEN t3.ITCLS IN ('WPLS') THEN 'Plastics' WHEN t3.ITCLS IN ('WVBC','WVCS') THEN 'Foundation' WHEN t3.ITCLS IN ('PANL') THEN 'Panel' WHEN t3.ITCLS IN ('ZKIZ') THEN 'ZipperCover' WHEN t3.ITCLS IN ('BBFR','WVHC') THEN 'Verona' WHEN t3.ITCLS NOT LIKE 'Z%' THEN 'RAW' ELSE 'Check' END) AS Product,
to_date(a.AAADAT||' '||right('000000'||ltrim(a.AAATIM),6), 'yyyymmdd hh24:mi:ss') as ScannedTime, 1/t2.ITMCQTY as Cartons
FROM  DISTLIBL.ACTAUDT a, (SELECT DISTINCT t4.ITNBR,t4.ITMCQTY FROM AFILELIBL.ITMEXT t4 GROUP BY t4.ITNBR,t4.ITMCQTY) AS t2, (SELECT  a2.ITNBR,a2.ITCLS FROM AMFLIBL.ITMRVA a2 WHERE a2.STID IN ('51')) AS t3
WHERE a.AAFARA IN ('RM') AND a.AACOD1 IN ('MV') and a.AASER#>0 AND a.AATARA LIKE 'HJ%' and CHAR(to_date(a.AAADAT||' '||right('000000'||ltrim(a.AAATIM),6),'yyyymmdd hh24:mi:ss')) BETWEEN ?  AND ? AND a.AAITM# = t2.ITNBR  and a.AAITM# = t3.ITNBR) AS X1
LEFT JOIN 
(SELECT CHAR(t1.AASER#) AS SN, t1.AAORD#
FROM DISTLIBL.ACTAUDT t1
WHERE t1.AACOD1 IN ('RM') AND t1.AACOD2 IN ('SN') AND t1.AATARA IN ('RM') AND t1.AAADAT Between CHAR(VARCHAR_FORMAT(current date -30 days,'YYYYMMDD'))  AND CHAR(VARCHAR_FORMAT(current date, 'YYYYMMDD')) 
) AS Y1 ON X1.SN = Y1.SN
ORDER BY X1.AAITM#, X1.ScannedTime