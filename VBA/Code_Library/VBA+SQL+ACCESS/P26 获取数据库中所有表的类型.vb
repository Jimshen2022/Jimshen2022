Option Explicit

Sub 获取数据库中所有表的名称与类型()
    
    Dim i As Integer
    Dim myData As String
    Dim cnn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    myData = ThisWorkbook.Path & "\学生管理.accdb"
    With cnn
         .Provider = "microsoft.ace.oledb.12.0"
         .Open myData
        
    End With
    Cells.Clear
    Range("a1:b1") = Array("表名称", "表类型")
    i = 2
    '开始获取表名称与类型
    Set rs = cnn.OpenSchema(adSchemaTables)
    Do Until rs.EOF
        If rs! table_type = "TABLE" Then
            Cells(i, 1) = rs("table_name")
            Cells(i, 2) = rs("table_type")
            i = i + 1
        End If
        rs.MoveNext
    Loop
    Columns.AutoFit
    rs.Close
    cnn.Close
    Set rs = Nothing
    Set cnn = Nothing
    
End Sub

