Sub WH335_TRX()
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
    
    
    Worksheets("TRX").Select
    Sheet3.Cells.Clear
    Set rs = New Recordset
    If rs.State = 1 Then rs.Close
    
    
    sql = "SELECT T1.TCODE, T1.ORDNO, T1.ITNBR, T2.ITCLS, T1.HOUSE, T1.UPDDT, T1.UPDTM, T1.TRQTY, T1.TRNDT, T1.LBHNO, T1.REFNO, T1.REASN, T1.USRSQ " &  _
            "FROM AMFLIBA.IMHIST T1, AMFLIBA.ITMRVA T2, AMFLIBA.WHSMST T3  " &  _
            "WHERE T2.ITNBR = T1.ITNBR AND T2.STID = T3.STID AND T1.HOUSE = T3.WHID AND T1.HOUSE='335' AND T1.UPDDT>=" & DateStart & " And T1.UPDDT<= " & DateEnd & " AND T1.TRQTY<>0 AND T1.TCODE in ('RP','SA') " &  _
            "ORDER BY T1.ITNBR "
    
    
    
    rs.Open sql, cnn, 3, 3
    For i = 0 To rs.Fields.Count - 1
        Worksheets("TRX").Cells(1, i + 1) = rs.Fields(i).Name
    Next i
    Sheet1.Columns("A:G").NumberFormat = "@"
    Worksheets("TRX").Range("A2").CopyFromRecordset rs
    Sheet1.Columns("A:G").AutoFit
    
    
    Application.ScreenUpdating = True
    'MsgBox "Data Downloaded Successful!"
    
    
End Sub
