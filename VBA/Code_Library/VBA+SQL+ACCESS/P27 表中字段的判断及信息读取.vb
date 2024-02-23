Option Explicit
Sub 检查字段是否存在()
    
    Dim myData As String
    Dim myTable As String
    Dim myColumn As String
    Dim cnn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    myData = ThisWorkbook.Path & "\学生管理.accdb"
    myTable = "学生"
    myColumn = "姓名"
    
    '建立数据库连接
    With cnn
         .Provider = "microsoft.ace.oledb.12.0"
         .Open myData
        
    End With
    
    '利用connection对象的openSchema方法产生字段记录集
    Set rs = cnn.OpenSchema(adSchemaColumns)
    
    '方法一
    '利用循环查询是否存在该数据表
    'Do While Not rs.EOF
    'rs!table_name也可以写成rs("table_name")
    'if Lcase(rs!table_name) = Lcase(myTable) then  'Lcase是将字母变小写
    'msgbox "数据表<" & myTable & ">存在。"
    'go to hhh
    'End If
    'rs.movenext
    'Loop
    'msgbox "数据表<" & myTable & ">不存在. "
    
    '方法二
    '利用recordset对象的Find方法查找数据表并判断是否存在
    'Find方法会直接将光标定位到找到的记录，如果没有找到，就定位到EOF.
    rs.Find "column_name='" & myColumn & "'"
    If rs.EOF Then
        MsgBox "数据表<" & myTable & ">不存在字段<" & myColumn & ">"
    Else
        MsgBox "数据表<" & myTable & ">存在字段<" & myColumn & ">"
    End If
hhh:
    rs.Close
    cnn.Close
    Set rs = Nothing
    Set cnn = Nothing
    
End Sub



Sub 获取字段名称类型及大小()
    Dim myData As String
    Dim cnn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim myTable As String
    myData = ThisWorkbook.Path & "\学生管理.accdb"
    myTable = "学生"
    
    With cnn
         .Provider = "microsoft.ace.oledb.12.0"
         .Open myData
        
    End With
    Cells.Clear
    
    Range("a1:c1") = Array("字段名", "字段类型", "字段大小")
    i = 2
    '开始获取表名称和表类型
    
    Dim myField As ADODB.Field
    rs.Open myTable, cnn, adOpenKeyset, adLockOptimistic
    For Each myField In rs.Fields
        Range("a" & i) = myField.Name
        'field.type用于获取字段的类型，但是不会直接返回类型的字符串
        '而是返回表示该类型的一个interger数字
        Range("b" & i) = myField. Type
        Range("c" & i) = myField.DefinedSize
        i = i + 1
    Next
    Columns.AutoFit
    rs.Close
    cnn.Close
    Set rs = Nothing
    Set cnn = Nothing
End Sub

