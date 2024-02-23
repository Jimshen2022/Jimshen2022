Option Explicit
'关于Connection 对象的openSchema方法
'格式： set recordset=connection.OpenSchema(查询类型)
'查询类型:
'adSchemaTables---数据表
'adSchemaColumns --字段
'adSchemaIndexes -- 索引
'adSchemaTables -- 数据表
'adSchemaTables --- 数据表
'adSchemaPrimaryKeys -- 主键

Sub 检查数据表是否存在()
    Dim myData As String
    Dim myTable As String
    Dim cnn As New ADODB.Connection
    Dim rs As ADODB.Recordset
    myData = ThisWorkbook.Path & "\成绩管理.accdb"
    myTable = "期末成绩"
    '建立数据库连接
    wiht cnn
    .Provider = "microsoft.ace.oledb.12.0"
    .Open myData
End With

'利用Connection对象的OpenSchema方法产生数据表记录集
Set rs = cnn.OpenSchema(adSchemaTables)

'利用循环查询是否存在该数据表
Do While Not rs.EOF
    If LCase(rs! table_name) = LCase(myTable) Then
        MsgBox "数据表<" & myTable & ">存在. "
        GoTo hhh
    End If
    rs.MoveNext
Loop
MsgBox "数据表<" & myTable & ">存在. "


'利用Recordset对象的Find方法查找数据表并判断是否存在
rs.Find "table_name='" & myTable & "'"
If rs.EOF Then
    MsgBox "数据表<" & myTable & ">不存在. "
Else
    MsgBox "数据表<" & myTable & ">存在。 "
End If

hhh:
rs.Close
cnn.Close
Set rs = Nothing
Set cnn = Nothing

End Sub