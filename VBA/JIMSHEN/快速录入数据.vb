'5.5.1 快速录入数据
Private Sub Worksheet_Change(ByVal Target As Range)
    '如果更改的单元各不是C列第3行以下的单元格或更改的单元格个数大于1时退出程序
    If Application.Intersect(Target, Range("c3:c65536")) Is Nothing Or Target.Count > 1 Then
        Exit Sub
    End If
    
    Dim i As Integer
    i = 3
    Do While Cells(i, "i").Value <> ""   ' 在参照表中循环
        '判断录入的字母与参照表的字母是否相符
        If UCase(Target.Value) = Cells(i, "I").Value Then
            Application.EnableEvents = False   '禁用事件,防止将字母改为商品名称时，再次执行该程序
                Target.Value = Cells(i, "I").Offset(0, 1).Value '写入产品信息
                Target.Offset(0, -1).Value = Date
                Target.Offset(0, 1).Value = Cells(i, "i").Offset(0, 2).Value   '写入商品代码
                Target.Offset(0, 2).Value = Cells(i, "i").Offset(0, 3).Value   '写入商品UP
                Target.Offset(0, 3).Select   '选中销售数量列，等待输入销售数量
            Application.EnableEvents = True
            Exit Sub
        End If
        i = i + 1
    Loop
End Sub
