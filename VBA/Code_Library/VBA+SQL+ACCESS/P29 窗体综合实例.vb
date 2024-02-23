
'公共变量
Option Explicit
Public cnn As ADODB.Connection
Public rs As ADODB.Recordset

Public Sub 数据连接()
    
    Set cnn = New ADODB.Connection
    With cnn
         .Provider = "microsoft.ace.oledb.12.0"
         .Open ThisWorkbook.Path & "\学生管理.accdb"
        
    End With
End Sub



'===================================================
'窗体中的code


Option Explicit
'加载窗体时，建立数据库连接，并刷新"数据表"列表框的信息

Private Sub UserForm_Initialize()
    '建立数据连接
    Call 数据连接
    
    '调用自定义过程，为"数据表清单"列表框刷新数据
    Call 获取数据表清单
End Sub

'自定义过程"获取数据表清单" 用于为数据表清单刷新数据

Public Sub 获取数据表清单()
    Set rs = cnn.OpenSchema(adSchemaTables)
    With 数据表清单
         .Clear
        Do Until rs.EOF
            If rs! table_type = "TABLE" Then
                 .AddItem rs! table_name
            End If
            rs.MoveNext
        Loop
         .ListStyle = fmListStyleOption '让每个选项都有按钮
    End With
    rs.Close
    Set rs = Nothing
End Sub

'单击“数据表清单”列表框，调用子过程，用于刷新所选表的字段列表
Private Sub 数据表清单_Click()
    Call 获取字段清单
End Sub

'子过程"获取字段清单" 用于获取所选表的字段，并显示在列表框中

Public Sub 获取字段清单()
    Dim sql As String, i As Integer
    '查询数据表，将字段名清单设置给"字段清单"列表框
    sql = "select * from " & 数据表清单.Text
    Set rs = New ADODB.Recordset
    rs.Open sql, cnn, adOpenKeyset, adLockOptimistic
    With 字段清单
         .Clear
        For i = 0 To rs.Fields.Count - 1
             .AddItem rs.Fields(i).Name
        Next i
         .ListStyle = fmListStyleOption '会在字段前显示可选圆点
    End With
    rs.Close
    Set rs = Nothing
End Sub

'单击"字段清单"列表框,调用子过程，将所选字段信息显示在文本框中

Private Sub 字段清单_Click()
    Call 获取字段信息
End Sub

'子过程 "获取字段信息" 用于获取所选字段的信息,并显示在文本框中
Public Sub 获取字段信息()
    Dim sql As String, i As Integer
    
    '查询选择的数据表
    sql = "select * from " & 数据表清单.Text
    Set rs = New ADODB.Recordset
    rs.Open sql, cnn, adOpenKeyset, adLockOptimistic
    
    '将字段的名称，类型，大小输出到对应的文本框
    字段名称.Value = rs.Fields(字段清单.Text).Name
    字段类型.Value = rs.Fields(字段清单.Text). Type
    字段大小.Value = rs.Fields(字段清单.Text).DefinedSize
    
End Sub


