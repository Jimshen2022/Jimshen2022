
Sub 内联结() '谭科VBA+Access P16
    Dim con As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    
    con.Open "provider=microsoft.ace.oledb.12.0;data source=" &  _
            ThisWorkbook.Path & "\学生管理.accdb"
    Dim sql As String
    
    '============================================================================================================
    '多表查询,内联结 ----查询所有课程的平均成绩，结果包含课程名称与平均成绩
    
    'sql = "select 课程名称,avg(成绩) as 平均成绩 " _
     & "from 课程 inner join 成绩 on 课程.课程代码=成绩.课程代码 " _
             & "group by 课程名称 having avg(成绩)>85"
    
    
    'sql执行的顺序:  From>where>group>having>select
    
    sql = "select 课程名称,avg(成绩) as 平均成绩 " _
             & "from 课程 inner join 成绩 on 课程.课程代码=成绩.课程代码 " _
             & "where 成绩>70 group by 课程名称"
    
    '============================================================================================================
    
    Set rs = con.Execute(sql) '执行sql命令，产生记录集
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

