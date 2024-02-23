Sub 对字段的增删改()
    
    Dim myData As String, myTable As String
    Dim cnn As New ADODB.Connection, sql As String
    
    myData = ThisWorkbook.Path & "\学生管理.accdb"
    myTable = "成绩"
    
    With cnn
         .Provider = "microsoft.ace.oledb.12.0"
         .Open myData
        
    End With
    
    '增加字段： alter table 表名 add 字段名 类型(大小)
    'sql = "alter table " & myTable & " add 数学 single, 备注 text(50)"
    
    '删除字段： alter table 表名 drop 字段
    'sql = "alter table 成绩 drop 数学"
    
    '修改字段类型与大小： alter table 表名 alter 字段 类型(大小）
    sql = "alter table 成绩 alter 课程代码 text(20)"
    
    ' On Error Resume Next
    '也可以用如下方式
    On Error GoTo hhh
    
    cnn.Execute sql
    cnn.Close
    Set cnn = Nothing
    Exit Sub
hhh:
    MsgBox Err.Description
    
End Sub
