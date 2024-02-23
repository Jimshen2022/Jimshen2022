Option Explicit
Dim cnn As ADODB.Connection
Dim myCmd As ADODB.Command
Dim rs As ADODB.Recordset

Sub 循环方式()
    '建立数据库连接
    Set cnn = New ADODB.Connection
    With cnn
         .Provider = "microsoft.ace.oledb.12.0"
         .Open ThisWorkbook.Path & "\成绩管理.accdb"
        
    End With
    
    '查询数据表是否已经存在
    Dim myTable As String
    myTable = "课程" '指定数据表名
    Set rs = cnn.OpenSchema(adSchemaTables)
    Do Until rs.EOF
        If LCase(rs! table_name) = LCase(myTable) Then
            GoTo hhh '如果存在则直接添加记录
        End If
        rs.MoveNext
    Loop
    
    '如果不存在,就创建数据表
    Set myCmd = New ADODB.Command
    Set myCmd.ActiveConnection = cnn
    myCmd.CommandText = "Create table " & myTable _
             & "(课程代码 text(20),课程名称 text(20),课程类别 text(8)," _
             & "学时 Integer,学分 integer,授课老师 text(10))"
    
    '利用command对象的execute 方法的执行命令
    myCmd.Execute, , adCmdText
    
hhh:
    
    Dim n%, i%, j%, sql$
    n = Range("a1").End(4).Row
    For i = 2 To n
        '检查是否存在某条记录
        sql = "select * from " & myTable _
                 & " where 课程代码 ='" & Cells(i, 1).Value & "'" _
                 & " and 课程名称='" & Cells(i, 2).Value & "'" _
                 & " and 课程类别='" & Cells(i, 3).Value & "'" _
                 & " and 学时=" & Cells(i, 4).Value _
                 & " and 学分=" & Cells(i, 5).Value _
                 & " and 授课老师='" & Cells(i, 6).Value & "'"
        
        Set rs = New ADODB.Recordset
        rs.Open sql, cnn, adOpenKeyset, adLockPessimistic
        If rs.RecordCount = 0 Then
            '如果数据表中没有工作表中某行数据，则添加数据
            rs.AddNew
            For j = 1 To rs.Fields.Count
                rs.Fields(j - 1) = Cells(i, j).Value
            Next j
            rs.Update
        Else
            '如果数据表中有工作表中某行数据，就将数据进行更新
            For j = 1 To rs.Fields.Count
                rs.Fields(j - 1) = Cells(i, j).Value
            Next j
            rs.Update
        End If
    Next i
    
    MsgBox "数据保存完毕. ", vbInformation, "提示"
    rs.Close
    cnn.Close
    Set rs = Nothing
    Set myCmd = Nothing
    Set cnn = Nothing
    
    
End Sub

Sub 数组方式()
    '将工作表数据存入数组
    Dim arr
    arr = Range("a1").CurrentRegion
    
    '建立数据库连接
    Set cnn = New ADODB.Connection
    With cnn
         .Provider = "microsoft.ace.oledb.12.0"
         .Open ThisWorkbook.Path & "\成绩管理.accdb"
    End With
    
    Dim myTable$, sql$
    myTable = "课程" '指定数据表名
    Dim i&, j&
    For i = 2 To UBound(arr)
        sql = "select * from " & myTable _
                 & " where 课程代码='" & arr(i, 1) & "'" _
                 & " and 课程名称='" & arr(i, 2) & "'" _
                 & " and 课程类别='" & arr(i, 3) & "'" _
                 & " and 学时=" & arr(i, 4) _
                 & " and 学分=" & arr(i, 5) _
                 & " and 授课教师='" & arr(i, 6) & "'"
        Set rs = New ADODB.Recordset
        rs.Open sql, cnn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount = 0 Then
            '如果数据表中没有工作表中某行数据，则添加数据
            rs.AddNew
            For j = 1 To rs.Fields.Count
                rs.Fields(j - 1) = arr(i, j)
            Next j
            rs.Update
        Else
            '如果数据表中有工作表中某行数据，就将数据进行更新
            For j = 1 To rs.Fields.Count
                rs.Fields(j - 1) = arr(i, j)
            Next j
            rs.Update
        End If
    Next i
    
    MsgBox "数据保存完毕. ", vbInformation, "提示"
    
    rs.Close
    cnn.Close
    Set rs = Nothing
    Set cnn = Nothing
    
End Sub


