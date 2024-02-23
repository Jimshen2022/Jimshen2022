'所谓前期绑定 ，是指在VBE中手工勾选引用Microsoft ADO相关类库 。
'在Excel中 ，按 < Alt + F11 > 快捷键打开VBA编辑窗口 ，依次单击 【工具 】→【引用 】，
'打开 【引用 - VBAProject 】对话框 。在 【可使用的引用 】列表框中 ，
'勾选 “Microsoft ActiveX Data Objects 2.8 Library ”库 ，或 “Microsoft ActiveX Data Objects 6.1 Library ”库 ，单击 【确定 】按钮关闭对话框 。



Sub 后期绑定()
    Dim cnn As Object
    Set cnn = CreateObject("adodb.connection")
End Sub

Sub Mycnn()
    Dim cnn As Object '定义变量
    Set cnn = CreateObject("adodb.connection") '后期绑定ADO
    cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties='Excel 12.0;HDR=yes;IMEX=0';Data Source=" & ThisWorkbook.FullName
    '建立链接
    cnn.Close '关闭链接
    Set cnn = Nothing '释放内存
End Sub


Sub Mycnn2()
    Dim cnn As Object
    Dim strPath As String
    Dim str_cnn As String
    Set cnn = CreateObject("adodb.connection")
    strPath = ThisWorkbook.FullName '当前工作簿的完整路径
    If Application.Version < 12 Then '判断Excel版本号，以使用不同的连接字符串
        str_cnn = "Provider=Microsoft.jet.OLEDB.4.0;Extended Properties=Excel 8.0;Data Source=" & strPath
    Else
        str_cnn = "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=Excel 12.0;Data Source=" & strPath
    End If
    cnn.Open str_cnn
    cnn.Close
    Set cnn = Nothing
End Sub


'这是一个最常用的VBA + ADO + SQL套路化查询代码 ，通常 ，我们只需要修改SQL语言 （第17行代码 ）以及放置查询结果的工作表名称 （第23行代码 ）

Sub DoSql_Execute1()
    Dim cnn As Object, rst As Object
    Dim strPath As String, str_cnn As String, strSQL As String
    Dim i As Long
    Set cnn = CreateObject("adodb.connection")
    '以上是第一步，后期绑定ADO
    '
    strPath = ThisWorkbook.FullName
    If Application.Version < 12 Then
        str_cnn = "Provider=Microsoft.jet.OLEDB.4.0;Extended Properties=Excel 8.0;Data Source=" & strPath
    Else
        str_cnn = "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=Excel 12.0;Data Source=" & strPath
    End If
    cnn.Open str_cnn
    '以上是第二步，建立链接
    '
    strSQL = "SELECT 姓名,成绩 FROM [Sheet1$] WHERE 成绩>=80"
    'Sql语句，查询Sheet1表成绩大于80……姓名和成绩的记录
    Set rst = cnn.Execute(strSQL)
    'Execute()执行SQL语句，始终得到一个新的记录集rst
    '以上是第三步，编写并执行SQL
    '
    Worksheets("结果表").Select '选中存放结果的工作表
    Cells.ClearContents '清空值
    For i = 0 To rst.Fields.Count - 1
        '利用fields属性获取所有字段名，fields包含了当前记录有关的所有字段,fields.count得到字段的数量
        '由于Fields.Count下标为0，又从0开始遍历，因此总数-1
        Cells(1, i + 1) = rst.Fields(i).Name
    Next
    Range("a2").CopyFromRecordset rst
    '使用单元格对象的CopyFromRecordset方法将rst内容复制到D2单元格为左上角的单元格区域
    '以上是第四步，将SQL查询结果和字段名写入表格指定区域
    '
    cnn.Close '关闭链接
    Set cnn = Nothing '释放内存
End Sub




