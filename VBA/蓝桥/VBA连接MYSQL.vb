Sub mysql_DB()
    
    Dim conn As ADODB.Connection     '定义连接对象
    Dim rs As ADODB.Recordset        '用于发送SQL语句
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    conn.connectionstring = "Driver={MySQL ODBC 8.0 Unicode Driver}; Server = LocalHost;DB=sql_jim;UID=root; PWD=1234"
    conn.Open  ' 连接数据库
    
    MsgBox ("连接成功! " & vbCrLf & "数据库状态: " & conn.State & vbCrLf & "数据库版本: " & conn.Version)
    
    conn.Close
    Set conn = Nothing
    Set rs = Nothing



End Sub




