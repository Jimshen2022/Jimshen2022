Sub cPullOHList()

t = Timer
Application.ScreenUpdating = False
Sheet5.Range("a1:k1048576").Cells.Clear

Set Db = New Connection
    Db.CursorLocation = adUseClient
    If Db.State = 1 Then Db.Close

'UserID = Sheet3.Range("a1").Value
'PW = Sheet3.Range("a2").Value

    Db.Open "Provider =IBMDASQL.DataSource.1" & _
     ";Catalog Library List=JDETSTDTA" & _
     ";Persist Security Info=True" & _
     ";Force Translate=0" & _
     ";Data Source = 10.9.3.115 " & _
     ";User ID =LLSEW1 " & _
     ";Password =LLSEW1 "

     Set adors = New Recordset
     If adors.State = 1 Then adors.Close

cmdtxt = "SELECT a1.ITNBR, a1.ITDSC, a1.ITCLS, a1.HOUSE, a1.LLOCN, a1.LQNTY, a1.ORDNO, a1.LBHNO, a2.RPAMVA AS "UnitPrice($USD)", a1.LQNTY*a2.RPAMVA AS "AMT($USD)", " & _
" (CASE WHEN a1.ITCLS LIKE 'TAF%' THEN 'RP' WHEN a1.ITCLS IN ('PACS') THEN 'UnKits' WHEN a1.ITCLS LIKE 'Z%' AND a1.ITCLS LIKE '%K' THEN 'UnKits' " & _
" WHEN a1.ITCLS IN ('ZACM','ZASU','ZMLH','ZMLR','ZUSR','ZUSU','ZVUC','ZXUC','ZUSU','ZUMU','ZAMU','ZASM','ZASR','ZDMA','ZMUC','ZSUS','ZUMS','ZUSM','ZVMA','ZVUS','ZXLH', " & _
" 'ZXLM','ZXLR','ZXMS','ZXMU') THEN 'UPH' " & _
      " WHEN a1.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB','ZDBC','ZABC','ZECD') THEN 'CG'  " & _
      " WHEN a1.ITCLS IN ('ZKIS') THEN 'Bedding'	 WHEN a1.ITCLS IN ('WPLS') THEN 'Plastics' WHEN a1.ITCLS IN ('WVBC','WVCS') THEN 'Foundation'	 " & _	
	  " WHEN a1.ITCLS IN ('PANL') THEN 'Panel' WHEN a1.ITCLS IN ('ZKIZ') THEN 'ZipperCover'  WHEN a1.ITCLS IN ('BBFR','WVHC') THEN 'Verona'	 " & _	
	  " WHEN a1.ITCLS NOT LIKE 'Z%' THEN 'RawMaterial' ELSE 'Check' END) AS Product " & _
" FROM (SELECT t1.ITNBR, t2.ITDSC, t2.ITCLS, t1.HOUSE, t1.LLOCN, t1.FDATE, t1.LQNTY, t1.ORDNO, t1.LBHNO   " & _
" FROM AMFLIBL.SLQNTY t1 left join AMFLIBL.ITMRVA t2 on t1.itnbr = t2.itnbr WHERE t1.LLOCN NOT IN ('S01ST1','S01PS1','FA00')) a1 " & _
" LEFT JOIN (SELECT b1.RPAITX, MAX(b1.RPAMVA) as RPAMVA FROM (SELECT x.RPAITX, x.ITCLS, x.RPAMVA, x.RPBLDT, x.RPZ0D7 " & _
" FROM (((SELECT a.RPAITX,(CASE WHEN a.RPBRCD IN ('VND') THEN a.RPAMVA/23090 ELSE a.RPAMVA END) AS RPAMVA,a.RPBLDT,a.RPZ0D7, T2.ITCLS FROM AMFLIBL.ITMFPR a  " & _
" LEFT JOIN AMFLIBL.ITMRVA T2 ON a.RPAITX=T2.ITNBR AND a.RPZ0D7 = T2.STID  " & _
" WHERE a.RPZ0D7 = '51' AND a.RPAITX||a.RPZ0D7||a.RPBLDT IN (SELECT a.RPAITX||a.RPZ0D7||MAX(a.RPBLDT) RPBLDT FROM AMFLIBL.ITMFPR a  WHERE a.RPZ0D7 = '51' GROUP BY a.RPAITX,a.RPZ0D7))  " & _
" UNION ALL  (SELECT t1.ITNO1G, t1.UCCT1G/23090 AS RPAMVA, t1.CCDT1G, t1.STID1G, t1.STID1G FROM AMFLIBL.ITMPRB t1)) " & _
" UNION ALL (SELECT t1.ITNBR, t1.LCOST/23090 AS RPAMVA, t1.LDQOH, t1.HOUSE, t1.ITCLS FROM AMFLIBL.ITEMBL t1))AS x " & _
" ORDER BY x.RPAITX, x.RPAMVA ASC) b1  GROUP BY b1.RPAITX ) a2  ON a1.ITNBR = a2.RPAITX "
	
	
	
    adors.Open cmdtxt, Db, 3, 3
		 For i = 0 To adors.Fields.Count - 1
			 Sheet5.Cells(1, i + 1) = adors.Fields(i).Name
		 Next i
    Sheet5.Activate
    sheet5.Columns("a:e").NumberFormat = "@"
    sheet5.Columns("g:h").NumberFormat = "@"	
    Sheet5.Range("a2").CopyFromRecordset adors
    adors.Close
    Set adors = Nothing
    
'    Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
'    Application.StatusBar = False
MsgBox Format(Timer - t, "0.00") & "s"
End Sub

