Sub MIL_OnHand()
    'On Error Resume Next

    Application.ScreenUpdating = False
    Dim i As Integer, j As Integer, n As Integer, m As Integer
    Dim sql As String
    Dim rs As New Recordset
    
    Set cnn = New Connection
    cnn.CursorLocation = adUseClient
    
    If cnn.State = 1 Then cnn.Close
    UName = Sheet4.Range("a1")
    UPass = Sheet4.Range("a2")
    DateStart = Sheet4.Range("a3")
    DateEnd = Sheet4.Range("a4")
    
    cnn.Open "Provider =IBMDASQL.DataSource.1" & _
            ";Catalog Library List=JDETSTDTA" & _
            ";Persist Security Info=True" & _
            ";Force Translate=0" & _
            ";Data Source = MILPROD" & _
            ";User ID =" & UName & "" & _
            ";Password =" & UPass
   
   
    Worksheets("OH").Select
    Sheet1.Cells.Clear
    Set rs = New Recordset
    If rs.State = 1 Then rs.Close
    
   
     sql = "SELECT t2.ITNBR, t1.ITDSC, t1.ITCLS, t2.HOUSE, t2.LLOCN, t2.FDATE, t2.LQNTY, t2.LBHNO,t3.ITMCQTY, " & _
            " (CASE " & _
                " WHEN t1.ITCLS LIKE 'TAF%' THEN 'RP' " & _
                " WHEN t1.ITCLS IN ('PACS') THEN 'UnKits' " & _
                " WHEN t1.ITCLS LIKE 'Z%' AND t1.ITCLS LIKE '%K' THEN 'UnKits' " & _
                " WHEN t1.ITCLS IN ('ZACM','ZASU','ZMLH','ZMLR','ZUSR','ZUSU','ZVUC','ZXUC','ZUSU','ZUMU') THEN 'UPH' " & _
                " WHEN t1.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB') THEN 'CG' " & _
                " WHEN t1.ITCLS IN ('ZKIS') THEN 'Bedding' " & _
                " WHEN t1.ITCLS IN ('WPLS') THEN 'Plastics' " & _
                " WHEN t1.ITCLS IN ('WVBC','WVCS') THEN 'Foundation' " & _
                " WHEN t1.ITCLS IN ('PANL') THEN 'Panel' " & _
                " WHEN t1.ITCLS IN ('ZKIZ') THEN 'ZipperCover' " & _
                " WHEN t1.ITCLS IN ('BBFR','WVHC') THEN 'Verona' " & _
                " WHEN t1.ITCLS NOT LIKE 'Z%' THEN 'RawMaterial' " & _
                " Else 'Check' END) AS Product " & _
            "  FROM AMFLIBL.ITEMASA t1, AMFLIBL.SLQNTY t2, AFILELIBL.ITMEXT t3 " & _
            "  WHERE t1.ITNBR = t2.ITNBR and t2.ITNBR = t3.ITNBR AND t2.HOUSE IN ('51','52') "
    
    
    rs.Open sql, cnn, 3, 3
    For i = 0 To rs.Fields.Count - 1
        Worksheets("OH").Cells(1, i + 1) = rs.Fields(i).Name
    Next i
   Sheet1.Columns("A:F").NumberFormat = "@"
   Worksheets("OH").Range("A2").CopyFromRecordset rs
   Sheet1.Columns("A:G").AutoFit
   
   
   Application.ScreenUpdating = True
   'MsgBox "Data Downloaded Successful!"
    
    
End Sub


Sub MIL_UP()
    'On Error Resume Next

    Application.ScreenUpdating = False
    't = Timer
    Dim i As Integer, j As Integer, n As Integer, m As Integer
    Dim sql As String
    Dim rs As New Recordset
    
    Set cnn = New Connection
    cnn.CursorLocation = adUseClient
    
    If cnn.State = 1 Then cnn.Close
    UName = Sheet4.Range("a1")
    UPass = Sheet4.Range("a2")
    DateStart = Sheet4.Range("a3")
    DateEnd = Sheet4.Range("a4")
    
    cnn.Open "Provider =IBMDASQL.DataSource.1" & _
            ";Catalog Library List=JDETSTDTA" & _
            ";Persist Security Info=True" & _
            ";Force Translate=0" & _
            ";Data Source = MILPROD" & _
            ";User ID =" & UName & "" & _
            ";Password =" & UPass
    
   
    Worksheets("UP").Select
    Sheet2.Cells.Clear
    Set rs = New Recordset
    If rs.State = 1 Then rs.Close
    
   'GET UP from AMFLIBL.ITMPRB and AMFLIBL.ITMFPR
    'sql = "SELECT x.RPAITX, x1.ITCLS, x1.RPAMVA, x1.RPBLDT, x1.RPZ0D7  "
     sql = "SELECT x.RPAITX, x.RPAMVA, x.RPBLDT, x.RPZ0D7, x.ITCLS  " & _
        "FROM (((SELECT a.RPAITX,(CASE WHEN a.RPBRCD IN ('VND') THEN a.RPAMVA/22715 ELSE a.RPAMVA END) AS RPAMVA,a.RPBLDT,a.RPZ0D7, T2.ITCLS  " & _
        "FROM AMFLIBL.ITMFPR a  " & _
        "LEFT JOIN AMFLIBL.ITMRVA T2 ON a.RPAITX=T2.ITNBR AND a.RPZ0D7 = T2.STID   " & _
        "WHERE a.RPZ0D7 = '51' AND a.RPAITX||a.RPZ0D7||a.RPBLDT IN (SELECT a.RPAITX||a.RPZ0D7||MAX(a.RPBLDT) RPBLDT FROM AMFLIBL.ITMFPR a  WHERE a.RPZ0D7 = '51' GROUP BY a.RPAITX,a.RPZ0D7))  " & _
        "UNION ALL " & _
        "(SELECT t1.ITNO1G, t1.UCCT1G/22715 AS RPAMVA, t1.CCDT1G, t1.STID1G, t1.STID1G FROM AMFLIBL.ITMPRB t1)) " & _
        "UNION ALL " & _
        "(SELECT t1.ITNBR, t1.LCOST/22715 AS RPAMVA, t1.LDQOH, t1.HOUSE, t1.ITCLS FROM AMFLIBL.ITEMBL t1))AS x  " & _
        "ORDER BY x.RPAITX, x.RPAMVA ASC"
    
    rs.Open sql, cnn, 3, 3
    For i = 0 To rs.Fields.Count - 1
        Worksheets("UP").Cells(1, i + 1) = rs.Fields(i).Name
    Next i
   
    Sheet2.Columns("A:A").NumberFormat = "@"
    Worksheets("UP").Range("A2").CopyFromRecordset rs
    Sheet2.Columns("A:B").AutoFit
   
   
   Application.ScreenUpdating = True
   'MsgBox "Data Downloaded Successful! total spend :   " & Format(Timer - t, "0.00" & "s")
    
    
End Sub

    
Sub MIL_TRX()
    'On Error Resume Next

    Application.ScreenUpdating = False
    't = Timer
    Dim i As Integer, j As Integer, n As Integer, m As Integer
    Dim sql As String
    Dim rs As New Recordset
    
    Set cnn = New Connection
    cnn.CursorLocation = adUseClient
    
    If cnn.State = 1 Then cnn.Close
    UName = Sheet4.Range("a1")
    UPass = Sheet4.Range("a2")
    DateStart = Sheet4.Range("a3")
    DateEnd = Sheet4.Range("a4")
    
    cnn.Open "Provider =IBMDASQL.DataSource.1" & _
            ";Catalog Library List=JDETSTDTA" & _
            ";Persist Security Info=True" & _
            ";Force Translate=0" & _
            ";Data Source = MILPROD" & _
            ";User ID =" & UName & "" & _
            ";Password =" & UPass
    
   
    Worksheets("TRX").Select
    Sheet3.Cells.Clear
    Set rs = New Recordset
    If rs.State = 1 Then rs.Close
    
   
     sql = "SELECT t1.TCODE, t1.ITNBR, t2.ITCLS, t1.HOUSE, WEEK(DATE('20'||SUBSTR(CHAR(T1.UPDDT),2,2)||'-'||SUBSTR(CHAR(T1.UPDDT),4,2)||'-'||SUBSTR(CHAR(T1.UPDDT),6,2))) AS WEEK,  " & _
           "'20'||SUBSTR(CHAR(T1.UPDDT),2,2)*100+WEEK(DATE('20'||SUBSTR(CHAR(T1.UPDDT),2,2)||'-'||SUBSTR(CHAR(T1.UPDDT),4,2)||'-'||SUBSTR(CHAR(T1.UPDDT),6,2))) AS YearWeek, SUM(t1.TRQTY) AS QTY " & _
           "FROM AMFLIBL.IMHIST t1, AMFLIBL.ITMRVA t2, AMFLIBL.WHSMST t3, AFILELIBL.ITMEXT t4 " & _
           "WHERE t2.ITNBR = t1.ITNBR and t1.itnbr=t4.itnbr AND t2.STID = t3.STID AND t1.HOUSE = t3.WHID AND t1.HOUSE IN ('51','52') " & _
           "AND t1.UPDDT >= " & DateStart & " And t1.UPDDT <= " & DateEnd & " AND t1.TRQTY<>0 AND t1.tcode IN ('IP','PQ','RM','SA','RP','RS','IS','IA','VR','IU','RC','SC','SM','SP','SS','LA')  " & _
            "GROUP BY  t1.TCODE, t1.ITNBR, t2.ITCLS, t1.HOUSE, WEEK(DATE('20'||SUBSTR(CHAR(T1.UPDDT),2,2)||'-'||SUBSTR(CHAR(T1.UPDDT),4,2)||'-'||SUBSTR(CHAR(T1.UPDDT),6,2))), " & _
            "'20'||SUBSTR(CHAR(T1.UPDDT),2,2)*100+WEEK(DATE('20'||SUBSTR(CHAR(T1.UPDDT),2,2)||'-'||SUBSTR(CHAR(T1.UPDDT),4,2)||'-'||SUBSTR(CHAR(T1.UPDDT),6,2)))  " & _
            "ORDER BY '20'||SUBSTR(CHAR(T1.UPDDT),2,2)*100+WEEK(DATE('20'||SUBSTR(CHAR(T1.UPDDT),2,2)||'-'||SUBSTR(CHAR(T1.UPDDT),4,2)||'-'||SUBSTR(CHAR(T1.UPDDT),6,2))) "
                

 
    rs.Open sql, cnn, 3, 3
    For i = 0 To rs.Fields.Count - 1
        Worksheets("TRX").Cells(1, i + 1) = rs.Fields(i).Name
    Next i
   Sheet1.Columns("A:D").NumberFormat = "@"
   Worksheets("TRX").Range("A2").CopyFromRecordset rs
   Sheet1.Columns("A:G").AutoFit
   
   
   Application.ScreenUpdating = True
   'MsgBox "Data Downloaded Successful!  " & Format(Timer - t, "0.00" & "s")
    
    
End Sub



























