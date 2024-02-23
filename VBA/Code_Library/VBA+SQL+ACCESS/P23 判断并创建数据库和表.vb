Sub P23() '判断并创建数据库和表
    
    Dim cnn As New ADODB.Connection
    Dim mycat As New ADOX.catalog
    Dim mydata As String
    Dim sql As String
    mydata = ThisWorkbook.Path & "\成绩管理.accdb" '指定数据库名称
    
    
    '利用Dir函数可以判断某个文件是否存在
    'dir(文件完整路径）
    '如果文件存在则返回文件名
    '如果文件不存在则返回空值
    'msgbox dir(mydata)
    
    If Len(Dir(mydata)) > 0 Then
        MsgBox "数据库已存在"
        Kill mydata '删除文件
    End If
    
    '创建数据库
    mycat.Create "provider=microsoft.ace.oledb.12.0;data source=" & mydata
    
    '连接数据库
    With cnn
         .Provider = "microsoft.ace.oledb.12.0"
         .Open mydata
    End With
    
    '创建数据表的SQL命令
    'create table 表名(字段 类型(宽度) 约束条件)
    sql = "create table 期中成绩(学号 text(10) not null," _
             & "姓名 text(8) not null,性别 text(1) not null," _
             & "班级 text(10) not null,语文 single not null," _
             & "数学 single not null,英语 single not null," _
             & "物理 single not null,化学 single not null," _
             & "生物 single not null,总分 single not null)"
    cnn.Execute sql
    MsgBox "database creation completed", vbInformation, "创建数据库"
    cnn.Close
    Set cnn = Nothing
    Set mycat = Nothing
    
End Sub
