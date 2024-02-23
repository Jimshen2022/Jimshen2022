

'RP 1020 vs AS400 variances report

'1.PULL DATA

Sub PullAS400OnHand()
    'On Error Resume Next
    Application.ScreenUpdating = False

    Dim i As Integer, j As Integer, n As Integer, m As Integer
    Dim cmdtxt As String
    Dim adors As New Recordset
    
    Set Db = New Connection
    Db.CursorLocation = adUseClient
    
    If Db.State = 1 Then Db.Close
    UName = Sheet1.Range("a1")
    UPass = Sheet1.Range("a2")
    DateStart = Sheet1.Range("a3")
    DateEnd = Sheet1.Range("a4")
    
    Db.Open "Provider =IBMDASQL.DataSource.1" & _
            ";Catalog Library List=JDETSTDTA" & _
            ";Persist Security Info=True" & _
            ";Force Translate=0" & _
            ";Data Source = 10.9.3.101 " & _
            ";User ID =" & UName & "" & _
            ";Password =" & UPass
    
   
    Worksheets("AS400vs1020").Select
    Sheet3.Cells.Clear
    Set adors = New Recordset
    If adors.State = 1 Then adors.Close

   
           cmdtxt = "SELECT a.ITNBR, a.HOUSE, a.ITCLS, a.MOHTQ, a.WHSLC, a.QTSYR, b.ITDSC " & _
            "FROM AMFLIBQ.ITEMBL a, AMFLIBQ.ITMRVA b, AMFLIBQ.WHSMST c " & _
            "WHERE b.ITCLS = a.ITCLS AND b.ITNBR = a.ITNBR AND a.HOUSE = c.WHID AND c.STID = b.STID AND ((a.MOHTQ<>0) AND (a.HOUSE='PC1' Or a.HOUSE='232')) AND (a.ITCLS NOT LIKE 'Z%') and a.itnbr <> 'SAMPLE' " & _
            "ORDER BY a.ITNBR "
   

    adors.Open cmdtxt, Db, 3, 3
    For i = 0 To adors.Fields.Count - 1
        Worksheets("AS400vs1020").Cells(1, i + 1) = adors.Fields(i).Name
    Next i
    
   Worksheets("AS400vs1020").Range("A2").CopyFromRecordset adors
   Sheet3.Columns("a:a").NumberFormat = "@"
   Sheet3.Columns("A:G").AutoFit
   
   
   Application.ScreenUpdating = True
   'MsgBox "Data Downloaded Successful!"
    
    
End Sub



Sub Pull102026nHand()
    'On Error Resume Next
    Application.ScreenUpdating = False

    Dim i As Integer, j As Integer, n As Integer, m As Integer
    Dim cmdtxt As String
    Dim adors As New Recordset
    
    Set Db = New Connection
    Db.CursorLocation = adUseClient
    
    If Db.State = 1 Then Db.Close
    UName = Sheet1.Range("a1")
    UPass = Sheet1.Range("a2")
    DateStart = Sheet1.Range("a3")
    DateEnd = Sheet1.Range("a4")
    
    Db.Open "Provider =IBMDASQL.DataSource.1" & _
            ";Catalog Library List=JDETSTDTA" & _
            ";Persist Security Info=True" & _
            ";Force Translate=0" & _
            ";Data Source = 10.9.3.101 " & _
            ";User ID =" & UName & "" & _
            ";Password =" & UPass
    
   
    Worksheets("1020.01.26").Select
    Sheet2.Cells.Clear
    Set adors = New Recordset
    If adors.State = 1 Then adors.Close

   
           cmdtxt = "SELECT a.LIWHSE, a.LIAREA, a.LIASLE, a.LISEC, a.LITIER, a.LIITEM, a.LIIDSC, a.LIIQTY, a.LIMDAT, a.LIMTME, a.LIMUSR, a.LIMPGM " & _
            "FROM DISTLIBQ.LOCINV a " & _
            "WHERE (a.LIWHSE='PC1') AND (a.LIIQTY<>0) and a.LIASLE <>1 " & _
            "ORDER BY a.LIITEM "
  
    adors.Open cmdtxt, Db, 3, 3
    For i = 0 To adors.Fields.Count - 1
        Worksheets("1020.01.26").Cells(1, i + 2) = adors.Fields(i).Name
    Next i
    
   Worksheets("1020.01.26").Range("b2").CopyFromRecordset adors
   Sheet2.Columns("a:a").NumberFormat = "@"
   Sheet2.Columns("A:G").AutoFit
   
   
   Application.ScreenUpdating = True
   'MsgBox "Data Downloaded Successful!"
    
    
End Sub



Sub PullASYard() 
    Dim wb As Workbook
    Dim arr, brr(), i&, j&, k&, nrow&, crr()
    
    't = Timer
    Application.ScreenUpdating = False
    'Sheet4.Range("a1").Value = "Data collected at:" & Format(Now(), "hhmm,mm-dd-yyyy")
    
    Sheet4.Activate
    Cells.Clear
    
    Set wb = GetObject("C:\Users\jishen\Downloads\ASYARD.xlsx") '´ò¿ª¹¤×÷²¾
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False
    ReDim brr(1 To UBound(arr), 1 To 18)
    For i = 1 To UBound(arr)
        For j = 1 To 18
            brr(i, j) = arr(i, j)
        Next
    Next
    
    With Sheet4
        .Columns("a:k").NumberFormat = "@"
        .Range("A1").Resize(UBound(arr), 18) = brr
        .Columns.AutoFit
    End With
    
    Erase arr
    Erase brr
    Application.ScreenUpdating = True
    'MsgBox Format(Timer - t, "0.00" & "s")
    
End Sub

Sub PullAdjustment()
    'On Error Resume Next
    Application.ScreenUpdating = False

    Dim i As Integer, j As Integer, n As Integer, m As Integer
    Dim cmdtxt As String
    Dim adors As New Recordset
    
    Set Db = New Connection
    Db.CursorLocation = adUseClient
    
    If Db.State = 1 Then Db.Close
    UName = Sheet1.Range("a1")
    UPass = Sheet1.Range("a2")
    DateStart = Sheet1.Range("a3")
    DateEnd = Sheet1.Range("a4")
    
    Db.Open "Provider =IBMDASQL.DataSource.1" & _
            ";Catalog Library List=JDETSTDTA" & _
            ";Persist Security Info=True" & _
            ";Force Translate=0" & _
            ";Data Source = 10.9.3.101 " & _
            ";User ID =" & UName & "" & _
            ";Password =" & UPass
    
   
    Worksheets("Adjustment").Select
    Sheet10.Cells.Clear
    Set adors = New Recordset
    If adors.State = 1 Then adors.Close

   
           cmdtxt = "SELECT a.LIWHSE, a.LIAREA, a.LIASLE, a.LISEC, a.LITIER, a.LIITEM, a.LIIDSC, a.LIIQTY, a.LIMDAT, a.LIMTME, a.LIMUSR, a.LIMPGM " & _
            "FROM DISTLIBQ.LOCINV a " & _
            "WHERE (a.LIWHSE='PC1') AND (a.LIIQTY<>0) and (a.LIASLE = 1) " & _
            "ORDER BY a.LIITEM "
  
    adors.Open cmdtxt, Db, 3, 3
    For i = 0 To adors.Fields.Count - 1
        Worksheets("Adjustment").Cells(1, i + 1) = adors.Fields(i).Name
    Next i
    
   Worksheets("Adjustment").Range("A2").CopyFromRecordset adors
   Sheet10.Columns("a:f").NumberFormat = "@"
   Sheet10.Columns("A:M").AutoFit
   
   
   Application.ScreenUpdating = True
   'MsgBox "Data Downloaded Successful!"
    
    
End Sub


