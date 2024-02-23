Sub 连接数据库()
    
    '给连接对象取名字
    'Dim con As ADODB.Connection
    
    '创建对象并赋值 (对象变量赋值前面必须加set)
    'Set con = New ADODB.Connection
    
    Dim con As New ADODB.Connection '这一句和上面2句的效果是一样的
    
    '建立联结
    '联结access
    'con.Open "provider=microsoft.ace.oledb.12.0;data source = " & ThisWorkbook.Path & "\学生管理.accdb"
    
    '另一种写法
    With con
         .Provider = "microsoft.ace.oledb.12.0"
         .ConnectionString = ThisWorkbook.Path & "\学生管理.accdb"
         .Open
    End With
    
    
    '联结access
    'con.Open "provider=microsoft.ace.oledb.12.0;extended properties=excel 12.0; data source = " & ThisWorkbook.Path & "\表一.xlsx"
    
    
    MsgBox "ok"
    
End Sub
