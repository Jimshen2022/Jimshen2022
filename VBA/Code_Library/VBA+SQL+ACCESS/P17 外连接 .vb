Sub 外连结()
    Dim con As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    con.Open "provider=microsoft.ace.oledb.12.0;data source=" & ThisWorkbook.Path & "\学生管理.accdb"
    Dim sql As String
    '----------------------------------------------------------------------------------------------------------
    '多表查询：外连接 --- from 左表 连接类型 右表 on 连接条件
    
    '左连接，左边连接字段有的，而右表没有的，左表全部显示，右表留空
    
    '例1 查询所有导师的院系信息，包含姓名，性别，职称，系号，系名
    'sql = "select 姓名,性别,职称,导师.院系编号,院系名 " _
     & "from 导师 left join 院系 on 导师.院系编号=院系.院系编号"
    
    
    
    '右连接，右表连接字段有的，而左表没有的，右表全部显示，左表留空。
    
    '例2 查询所有院系的导师信息，包含系号，系名，姓名，职称
    
    'sql = "select 院系.院系编号,院系名,姓名,职称 " _
     & "from 院系 right join 导师 on 导师.院系编号=院系.院系编号"
    
    
    '全连接 excel中不适用
    '例3 查询所有导师，所有院系的信息，包含姓名，性别，职称，系号，系名
    
    sql = "select 院系.院系编号,院系名,姓名,职称 " _
             & "from 院系 full join 导师 on 导师.院系编号=院系.院系编号"
    
    
    '----------------------------------------------------------------------------------------------------------
    Set rs = con.Execute(sql)
    Cells.Clear
    Dim i As Integer
    For i = 0 To rs.Fields.Count - 1
        Cells(1, i + 1) = rs.Fields(i).Name
    Next
    Range("a2").CopyFromRecordset rs
    Columns.AutoFit
    rs.Close: Set rs = Nothing
    con.Close: Set con = Nothing
    
End Sub
