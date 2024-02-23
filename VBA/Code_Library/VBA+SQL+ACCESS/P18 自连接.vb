Sub 自连接()
    
    Dim con As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    
    con.Open "provider=microsoft.ace.oledb.12.0;data source=" & ThisWorkbook.Path & "\学生管理.accdb"
    
    Dim sql$
    '-------------------------------------------------------------------------------------------------------
    ' query the employee table that have duplicated name
    ' query all data
    'sql = "select * from 员工 t1 inner join 员工 t2 on t1.姓名= t2.姓名 "
    
    ' 查询表中重复的姓名，但是员工编号不同的记录
    
    sql = "select distinct t1.编号,t1.姓名,t1.年龄,t1.职务,t1.部门 " _
             & "from 员工 t1 inner join 员工 t2 " _
             & "on t1.姓名=t2.姓名 where t1.编号<>t2.编号 order by t1.姓名"
    
    '-------------------------------------------------------------------------------------------------------
    
    Set rs = con.Execute(sql)
    Cells.Clear
    
    Dim i%
    
    For i = 0 To rs.Fields.Count - 1
        Cells(1, i + 1).Value = rs.Fields(i).Name
    Next
    Range("a2").CopyFromRecordset rs
    Columns.AutoFit
    rs.Close: Set rs = Nothing
    con.Close: Set con = Nothing
    
End Sub
