Sub WH335_UP()
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
    
    
    Worksheets("UP").Select
    Sheet2.Cells.Clear
    Set rs = New Recordset
    If rs.State = 1 Then rs.Close
    
    
    sql = "SELECT PRICE.PITEM, PRICE.PAMNT " &  _
            "FROM AFILELIB.PRICE PRICE  " &  _
            "WHERE  (PRICE.PRICCD='FOBARC') " &  _
            "ORDER BY PRICE.PITEM "
    
    
    rs.Open sql, cnn, 3, 3
    For i = 0 To rs.Fields.Count - 1
        Worksheets("UP").Cells(1, i + 1) = rs.Fields(i).Name
    Next i
    
    Sheet2.Columns("A:A").NumberFormat = "@"
    Worksheets("UP").Range("A2").CopyFromRecordset rs
    Sheet2.Columns("A:B").AutoFit
    
    
    Application.ScreenUpdating = True
    'MsgBox "Data Downloaded Successful!"
    
    
End Sub