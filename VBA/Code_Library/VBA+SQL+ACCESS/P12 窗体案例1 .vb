Sub 打开窗口() ‘在VBE forms下新增窗体frmEmpInfo ，并在module 写入语句 ，显示窗体的名字
    frmEmpInfo.Show
    
End Sub



'申明公共变量，因为不能查询中途断了连接
Dim con As ADODB.Connection '声明连接对象变量
Dim rs As ADODB.Recordset '声明记录集对象变量
Private Sub lstEmp_Click()
    Dim arr, i%
    Dim sql As String
    sql = "select *  from 员工 where 编号 ='" & Left(lstEmp.Value, 6) & "'"
    rs.Open sql, con, adOpenKeyset, adLockOptimistic
    '将每个字段的值存到控件中
    arr = Array("txtID", "txtName", "txtAge", "txtNumber", "txtDate", "txtAddress", "txtBM", "txtZW", "txtMail", "txtInfo")
    For i = 0 To UBound(arr)
        Me.Controls(arr(i)).Value = rs.Fields(i)
    Next i
    rs.Close
End Sub


'当窗体加载时，填写lstBM这个列表框的内容
Private Sub UserForm_Initialize()
    '建立数据库的连接
    Set con = New ADODB.Connection
    With con
         .Provider = "microsoft.ace.oledb.12.0"
         .ConnectionString = "data source = " & ThisWorkbook.Path & "\学生管理.accdb"
         .Open
        
    End With
    
    
    '提取不重复的部门名称
    Dim sql As String '定义命令字符串变量
    
    sql = "select distinct 部门 from 员工"
    Set rs = New ADODB.Recordset '创建记录集对象
    rs.Open sql, con, adOpenKeyset, adLockOptimistic
    '此时部门list已放到内存中
    
    Dim i%
    '将记录集中的部门名称显示到lstBM列表框中
    With lstBM
         .Clear '这个很重要，防止再次查询时，将上一次的记录仍保留着
        For i = 1 To rs.RecordCount '从1到记录集的个数进行循环
             .AddItem rs("部门")
            rs.MoveNext '将记录集中的指针指向下一条记录
        Next i
    End With
    
    rs.Close
    
End Sub


'关闭按钮---------释放变量空间，关闭数据库连接，关闭窗体
Private Sub cmdClose_Click()
    con.Close
    Set rs = Nothing
    Set con = Nothing
    Unload Me '如果不加此句，点关闭按钮后，再次点击会报错，因为rs,con都关闭了，再让它去关闭 所以会报错。
    
End Sub


'鼠标选择某个部门，相当于单击列表框，单击列表框，查询所选部门的员工
'提取员工的编号与姓名，避免姓名重复的问题
Private Sub LstBM_Click()
    Dim sql As String, i As Integer
    sql = "select 编号,姓名 from 员工 where 部门= '" & lstBM.Value & " ' order by 编号 "
    rs.Open sql, con, adOpenKeyset, adLockOptimistic
    With lstEmp
         .Clear '这里的clear很重要，若不加会出现点一次部门就累加一次编号与姓名
        For i = 1 To rs.RecordCount
             .AddItem rs("编号") & Space(2) & rs("姓名")
            rs.MoveNext
        Next i
    End With
    rs.Close
End Sub


