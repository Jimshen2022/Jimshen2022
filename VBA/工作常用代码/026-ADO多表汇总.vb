Sub 如何使用VBA进行多表汇总()
    Dim AdoConn As New ADODB.Connection
    Dim AdoRst As ADODB.Recordset
    Dim strConn As String
    Dim strSQL As String
    Application.ScreenUpdating = False
    '设置连接字符串
    strConn = " Provider=Microsoft.ACE.OLEDB.12.0;" &  _
            "Data Source=" & ThisWorkbook.FullName &  _
            ";Extended Properties=""Excel 12.0;HDR=YES"";"
    '打开数据库链接
    AdoConn.Open strConn
    '获取数据库表格结构
    Set AdoRst = AdoConn.OpenSchema(adSchemaTables)
    '获取各表名称，编写SQL语句
    Do Until AdoRst.EOF
        
        If AdoRst! TABLE_TYPE = "TABLE" And AdoRst! TABLE_NAME <> "汇总表$" And AdoRst! TABLE_NAME Like "'*$'" Then
            strSQL = strSQL & "Union All Select * From [" & AdoRst! TABLE_NAME & "] "
        End If
        AdoRst.MoveNext
    Loop
    '关闭数据记录集
    AdoRst.Close
    '重新编辑SQL语句，去掉最开头的Union All，共10个字符
    strSQL = Right(strSQL, Len(strSQL) - 10)
    '编辑SQL语句，计算应发合计的总合
    strSQL = "Select 员工编号,姓名,Sum(应发合计) as 全年应发合计 From (" & strSQL & ") Group By 员工编号,姓名"
    '执行查询
    Set AdoRst = AdoConn.Execute(strSQL)
    '将结果写入汇总表
    With Sheet1
         .Cells.Clear
         .Range("A2").CopyFromRecordset AdoRst
        '填写标题
        For i = 1 To AdoRst.Fields.Count
             .Cells(1, i) = AdoRst.Fields(i - 1).Name
        Next
        '自动调整列宽
         .UsedRange.Columns.AutoFit
        '设置边框颜色
         .UsedRange.Borders.Color = 0
        '设置标题行颜色
         .UsedRange.Rows(1).Style = "强调文字颜色 2"
    End With
    AdoRst.Close
    '关闭数据库连接
    AdoConn.Close
    Application.ScreenUpdating = True
End Sub