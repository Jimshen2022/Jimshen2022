Sub cPullSQLData() ' faster

t = Timer
Application.ScreenUpdating = False
'Application.Calculation = xlCalculationManual
'Application.StatusBar = "Loading UOM, please wait ......"

Worksheets("DATA").Cells.Clear
Range("s1") = "DataCollectedAt:  " & Format(Now, "HH:MM:SS am/pm,  mmm.dd.yyyy")
Range("s1").Font.Color = -16776961

Set Db = New Connection
    Db.CursorLocation = adUseClient
    If Db.State = 1 Then Db.Close

UserID = Sheet3.Range("a1").Value
PW = Sheet3.Range("a2").Value

    Db.Open "Provider =IBMDASQL.DataSource.1" & _
     ";Catalog Library List=JDETSTDTA" & _
     ";Persist Security Info=True" & _
     ";Force Translate=0" & _
     ";Data Source = 10.9.3.115 " & _
     ";User ID = " & UserID & "" & _
     ";Password = " & PW

     Set adors = New Recordset
     If adors.State = 1 Then adors.Close

    cmdtxt = "Select Y1.SN,Y1.TDTSTS,Y1.TDITEM,Y1.MO,Y1.TDWHSE,Y1.TDMDAT,Y1.TDMTME,Y1.TXT_TIME,Y1.ITCLS,(1/Y1.ITMCQTY) as Cartons, Y1.Product,Y1.SHIFT,Y1.Line,Y2.CTN,Y2.""CTN_Status"" " & _
            "From(SELECT X1.SN,X1.TDTSTS,X1.TDITEM,X1.MO,X1.TDWHSE,X1.TDMDAT,X1.TDMTME,X1.TXT_TIME,X1.ITCLS,X1.Product, X1.SHIFT,CHAR(TRIM(SUBSTR(X2.JOBNO,1,5))) AS Line,X1.ITMCQTY " & _
            "FROM(SELECT CHAR(trim(t1.TDTAG#)) AS SN,t1.TDITEM,t1.TDAPO# as MO,t1.TDWHSE,t1.TDMDAT,t1.TDMTME, t1.TDTSTS,right('000000'||ltrim(t1.TDMTME),6) AS TXT_TIME,t3.ITCLS,t5.ITMCQTY, " & _
            "(CASE WHEN t3.ITCLS IN ('WPLS') THEN 'Plastics' WHEN t3.ITCLS IN ('WVBC','WVHC') THEN 'Foundation' " & _
                "WHEN t3.ITCLS IN ('SLDK') THEN 'RP' WHEN t3.ITCLS LIKE 'T%' THEN 'RP' WHEN t3.ITCLS IN ('ZKIS') THEN 'Bedding' WHEN t3.ITCLS IN ('ZKIZ') THEN 'ZipperCover' " & _
                "WHEN t3.ITCLS LIKE 'Z%' AND t3.ITCLS LIKE '%K' THEN 'UnKits' WHEN t3.ITCLS IN ('PACS') THEN 'UnKits' WHEN t3.ITCLS IN ('BBFR') THEN 'FR SOCK' WHEN t3.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB') THEN 'CG' " & _
                "WHEN t3.ITCLS LIKE 'Z%' THEN 'UPH' ELSE 'Check' END) AS Product, (CASE WHEN t1.TDMTME BETWEEN '070000' AND '194459' THEN 'DS' ELSE 'NS' END) AS SHIFT " & _
            "FROM DISTLIBL.TAGINVD t1,(SELECT DISTINCT t2.ITNBR,t2.ITCLS FROM AMFLIBL.ITEMBL t2 WHERE t2.HOUSE = '51' GROUP BY t2.ITNBR,t2.ITCLS) AS t3, " & _
            "(SELECT DISTINCT t4.ITNBR,t4.ITMCQTY FROM AFILELIBL.ITMEXT t4 GROUP BY t4.ITNBR,t4.ITMCQTY) AS t5 " & _
            "WHERE t1.TDITEM = t3.ITNBR AND t1.TDITEM=t5.ITNBR AND t3.ITNBR=t5.ITNBR and t1.TDTSTS IN ('R','S') and t1.TDMDAT " & _
            "BETWEEN  int('1'||substr(trim(char(CURRENT DATE - 60 days)),3,2)||substr(trim(char(CURRENT DATE- 60 days)),6,2)||substr(trim(char(CURRENT DATE- 60 days)),9,2))  " & _
            "AND int('1'||substr(trim(char(CURRENT DATE + 1 days)),3,2)||substr(trim(char(CURRENT DATE + 1 days)),6,2)||substr(trim(char(CURRENT DATE + 1 days)),9,2)) " & _
            "AND NOT EXISTS (SELECT 1 FROM (SELECT CHAR(trim(a.WCSSERIALNUMBER)) AS SN  FROM LLUSAF.WVCNTSD a, LLUSAF.WVCNTHD b WHERE a.WCSCONTAINERNUMBER = b.WCHCONTAINERNUMBER and b.WCHCONTAINERSTATUS IN ('P','T') AND b.WCHPOSTEDTIMESTAMP BETWEEN CHAR(CURRENT DATE - 61 days) AND CHAR(CURRENT DATE + 1 days)) XX1 WHERE CHAR(trim(t1.TDTAG#)) = XX1.SN) " & _
            "AND NOT EXISTS (SELECT 1 FROM (SELECT CHAR(TRIM(a.WCSSERIALNUMBER)) AS SN  FROM ASHLEYARCL.WVCNTSDA a  WHERE a.WCSADDEDTIMESTAMP between char(current date - 61 days) and char(current DATE + 1 days)) XX2 WHERE CHAR(trim(t1.TDTAG#)) = XX2.SN)) AS X1 " & _
            "LEFT JOIN ((SELECT t1.ORDNO,t1.JOBNO,t1.FITEM FROM AMFLIBL.MOMAST t1 where t1.CRDT " & _
            "BETWEEN  int('1'||substr(trim(char(CURRENT DATE - 180 days)),3,2)||substr(trim(char(CURRENT DATE- 180 days)),6,2)||substr(trim(char(CURRENT DATE- 180 days)),9,2)) " & _
            "AND int('1'||substr(trim(char(CURRENT DATE + 1 days)),3,2)||substr(trim(char(CURRENT DATE + 1 days)),6,2)||substr(trim(char(CURRENT DATE + 1 days)),9,2)) " & _
            "order by t1.CRDT DESC) UNION ALL (SELECT t1.ORDNO,t1.JOBNO,t1.FITEM FROM AMFLIBL.MOHMST t1 where t1.CRDT " & _
            "BETWEEN  int('1'||substr(trim(char(CURRENT DATE - 180 days)),3,2)||substr(trim(char(CURRENT DATE- 180 days)),6,2)||substr(trim(char(CURRENT DATE- 180 days)),9,2))  " & _
            "AND int('1'||substr(trim(char(CURRENT DATE + 1 days)),3,2)||substr(trim(char(CURRENT DATE + 1 days)),6,2)||substr(trim(char(CURRENT DATE + 1 days)),9,2)) " & _
            "order by t1.CRDT DESC)) AS X2 ON X1.MO=X2.ORDNO and X1.TDITEM=X2.FITEM) as Y1 Left join (SELECT CHAR(trim(a.WCSSERIALNUMBER)) AS SN,a.WCSCONTAINERNUMBER AS CTN, " & _
            "(CASE WHEN a.WCSCONTAINERNUMBER LIKE 'MRUN%' THEN 'InTempCTN' WHEN a.WCSCONTAINERNUMBER LIKE 'KECR%' THEN 'InTempCTN' WHEN a.WCSCONTAINERNUMBER LIKE 'KHO%' THEN 'InTempCTN' WHEN a.WCSCONTAINERNUMBER LIKE 'M3K%' THEN 'InTempCTN' " & _
            "WHEN a.WCSCONTAINERNUMBER LIKE 'M3E%' THEN 'InTempCTN' WHEN a.WCSCONTAINERNUMBER LIKE 'M3H%' THEN 'InTempCTN' WHEN a.WCSCONTAINERNUMBER LIKE 'RUN%' THEN 'InTempCTN' ELSE 'InRealCTN' END) AS ""CTN_Status"" FROM LLUSAF.WVCNTSD a, LLUSAF.WVCNTHD b  " & _
            "WHERE a.WCSCONTAINERNUMBER = b.WCHCONTAINERNUMBER and b.WCHCONTAINERSTATUS NOT IN ('P','T')) AS Y2 ON Y1.SN = Y2.SN limit 1000 "
            
            
    
    adors.Open cmdtxt, Db, 3, 3
     For i = 0 To adors.Fields.Count - 1
         Sheet2.Cells(1, i + 1) = adors.Fields(i).Name
     Next i

    Sheet2.Activate
    Columns("a:o").NumberFormat = "@"

     Sheet2.Range("a2").CopyFromRecordset adors
     adors.Close
     Set adors = Nothing
    Range("x1") = "Finished At:  " & Format(Now, "HH:MM:SS am/pm,  mmm.dd.yyyy")

    ActiveWorkbook.Save
    
'    Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
'    Application.StatusBar = False
MsgBox Format(Timer - t, "0.00") & "s"


End Sub






