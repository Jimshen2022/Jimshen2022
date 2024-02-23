   
Sub MIL_TRX()
    'On Error Resume Next

    Application.ScreenUpdating = False
    t = Timer
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
    
   
    Worksheets("TRX").Activate
    Sheet3.Cells.Clear
    Set rs = New Recordset
    If rs.State = 1 Then rs.Close
    
   
     sql = "SELECT t1.HOUSE,t1.TCODE,t1.ORDNO,t1.ITNBR,t2.ITCLS, t1.UPDDT,t1.UPDTM,t1.TRQTY,t1.ENTUM,t1.VNDNR,t1.REFNO,t1.LLOCN,t1.BATCH,t1.TRMID, " & _
           "CHAR(t1.UPDDT||' '||right('000000'||ltrim(t1.UPDTM),6)) AS TrxTime, CHAR(SUBSTR(right('000000'||ltrim(t1.UPDTM),6),1,2)) AS HOUR, " & _
		   "(CASE WHEN t2.ITCLS IN ('SLDK') THEN 'RP' " & _
				"WHEN t2.ITCLS LIKE 'T%' THEN 'RP' " & _
				"WHEN t2.ITCLS LIKE 'R%' THEN 'RP' " & _
				"WHEN t2.ITCLS IN ('PACS') THEN 'UnKits' " & _
				"WHEN t2.ITCLS LIKE 'Z%' AND t2.ITCLS LIKE '%K' THEN 'UnKits' " & _
				"WHEN t2.ITCLS IN ('ZACM','ZASU','ZMLH','ZMLR','ZUSR','ZUSU','ZVUC','ZXUC','ZUSU','ZUMU') THEN 'UPH' " & _
				"WHEN t2.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB') THEN 'CG' " & _
				"WHEN t2.ITCLS IN ('ZKIS') THEN 'Bedding' " & _		
				"WHEN t2.ITCLS IN ('WPLS') THEN 'Plastics' " & _
				"WHEN t2.ITCLS IN ('WVBC','WVCS') THEN 'Foundation' " & _		
				"WHEN t2.ITCLS IN ('PANL') THEN 'Panel' " & _
				"WHEN t2.ITCLS IN ('ZKIZ') THEN 'ZipperCover' " & _
				"WHEN t2.ITCLS IN ('BBFR','WVHC') THEN 'Verona'	" & _	 
				"WHEN t2.ITCLS NOT LIKE 'Z%' THEN 'RawMaterial' " & _
				"ELSE 'Check' END) AS Product " & _	
           "FROM AMFLIBL.IMHIST t1, AMFLIBL.ITMRVA t2, AMFLIBL.WHSMST t3 " & _
           "WHERE t1.ITNBR=t2.ITNBR AND t2.STID = t3.STID AND t1.HOUSE = t3.WHID AND t1.TRQTY > 0 AND t1.TCODE IN ('RP','RM','PQ')  " & _
           "AND CHAR(t1.UPDDT||' '||right('000000'||ltrim(t1.UPDTM),6)) BETWEEN " & DateStart & " And " & DateEnd & _
           "AND t2.ITCLS NOT LIKE 'Z%' "

 
    rs.Open sql, cnn, 3, 3
    For i = 0 To rs.Fields.Count - 1
        Worksheets("TRX").Cells(1, i + 1) = rs.Fields(i).Name
    Next i
   Sheet1.Columns("A:F").NumberFormat = "@"
   Sheet1.Columns("H:O").NumberFormat = "@"
   
   Worksheets("TRX").Range("A2").CopyFromRecordset rs
   Sheet1.Columns("A:O").AutoFit
   
   
   Application.ScreenUpdating = True
   MsgBox "Data Downloaded Successful!  " & Format(Timer - t, "0.00" & "s")
    
    
End Sub


