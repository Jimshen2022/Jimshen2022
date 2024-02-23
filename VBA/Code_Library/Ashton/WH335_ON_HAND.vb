Sub WH335_OnHand()
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
    
    cnn.Open "Provider =IBMDASQL.DataSource.1" &  _
            ";Catalog Library List=JDETSTDTA" &  _
            ";Persist Security Info=True" &  _
            ";Force Translate=0" &  _
            ";Data Source = AFIPROD" &  _
            ";User ID =" & UName & "" &  _
            ";Password =" & UPass
    
    
    Worksheets("DATA").Select
    Sheet1.Cells.Clear
    Set rs = New Recordset
    If rs.State = 1 Then rs.Close
    
    
    sql = "SELECT T1.ITNBR, T1.HOUSE, T1.ITCLS, T1.MOHTQ, T1.WHSLC, T1.QTSYR, T2.ITDSC " &  _
            "FROM AMFLIBA.ITEMBL T1, AMFLIBA.ITMRVA T2, AMFLIBA.WHSMST T3  " &  _
            "WHERE  T2.ITCLS = T1.ITCLS AND T2.ITNBR = T1.ITNBR AND T2.STID = T3.STID AND T1.HOUSE = T3.WHID AND ((T1.HOUSE='335') AND (T1.MOHTQ<>0)) " &  _
            "ORDER BY T1.ITNBR "
    
    
    rs.Open sql, cnn, 3, 3
    For i = 0 To rs.Fields.Count - 1
        Worksheets("DATA").Cells(1, i + 1) = rs.Fields(i).Name
    Next i
    
    Worksheets("DATA").Range("A2").CopyFromRecordset rs
    Sheet1.Columns("A:C").NumberFormat = "@"
    Sheet1.Columns("A:G").AutoFit
    
    
    Application.ScreenUpdating = True
    'MsgBox "Data Downloaded Successful!"
    
    
End Sub