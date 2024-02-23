'方法一
Sub 利用command对象创建表()
    
    On Error Resume Next
    Dim myCmd As New ADODB.Command
    Dim myCat As New ADOX.Catalog
    Dim mydata As String
    Dim myTable As String
    Dim sql As String
    mydata = ThisWorkbook.Path & "\成绩管理.accdb"
    myTable = "期末成绩"
    '建立数据库连接
    myCat.ActiveConnection = "Provider=microsoft.ace.oledb.12.0;" _
             & "data source=" & mydata
    '删除已存在的同名数据库
    myCat.Tables.Delete myTable
    
    '设置数据库连接
    Set myCmd.ActiveConnection = myCat.ActiveConnection
    '设置创建数据表的SQL语句
    sql = "create table 期末成绩(学号 text(10) not null," _
             & "姓名 text(8) not null,性别 text(1) not null," _
             & "班级 text(10) not null,语文 single not null," _
             & "数学 single not null,英语 single not null," _
             & "物理 single not null,化学 single not null," _
             & "生物 single not null,总分 single not null)"
    
    '利用Command对象的Execute方法执行命令
    With myCmd
         .CommandText = sql
         .Execute, , adCmdText '表示执行一个文本命令
    End With
    MsgBox "数据表创建成功!", vbInformation, "创建数据表"
    Set myCmd = Nothing
    Set myCat = Nothing
    
End Sub


'方法二
Sub 利用sql语句创建数据表()
    Dim cnn As New ADODB.Connection
    Dim mydata As String
    Dim myTable As String
    Dim sql As String
    mydata = ThisWorkbook.Path & "\成绩管理.accdb"
    myTable = "期末成绩"
    
    '建立数据库连接
    With cnn
         .Provider = "microsoft.ace.oledb.12.0"
         .Open mydata
    End With
    
    '删除已有的同名数据表
    sql = "drop table " & myTable
    cnn.Execute sql
    
    '设置创建数据表的SQL语句
    sql = "create table 期末成绩(学号 text(10) not null," _
             & "姓名 text(8) not null,性别 text(1) not null," _
             & "班级 text(10) not null,语文 single not null," _
             & "数学 single not null,英语 single not null," _
             & "物理 single not null,化学 single not null," _
             & "生物 single not null,总分 single not null)"
    
    '利用connection对象的Execute方法执行命令
    cnn.Execute sql
    cnn.Close
    Set cnn = Nothing
    MsgBox "数据表创建成功!", vbInformation, "创建数据表"
    
End Sub

