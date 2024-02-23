Sub Dicttl1()
    Dim d As Object, arr, brr, i&
    Set d = CreateObject("scripting.dictionary")
    '后期字典
    'd.CompareMode = vbTextCompare
    '不区分字母大小写
    arr = Range("a1:b" & Cells(Rows.Count, 1).End(xlUp).Row)
    '数据源装入数组arr
    For i = 1 To UBound(arr)
    '遍历数据源，累加姓名成绩
        d(arr(i, 1)) = d(arr(i, 1)) + Val(arr(i, 2))
        'val函数提取纯数值，如果是纯文本值则计算为0，避免文本值数学运算出错
        '如果是重复值计数，可以改成如下：
        'd(arr(i, 1)) = d(arr(i, 1)) + 1
    Next
    brr = Range("d1:f" & Cells(Rows.Count, 4).End(xlUp).Row)
    '查询区域装入数组brr
    For i = 2 To UBound(brr)
        If d.exists(brr(i, 1)) Then
        '如果字典中存在查询的姓名,则提取总成绩
            brr(i, 3) = d(brr(i, 1))
        Else
        '否则返回空文本
            brr(i, 3) = ""
        End If
    Next
    With Range("d1:f" & Cells(Rows.Count, 4).End(xlUp).Row)
        .NumberFormat = "@" '设置文本格式，防止某些文本数值数据变形
        .Value = brr
        'brr数组放回单元格区域
    End With
    Set d = Nothing
    '释放字典
    MsgBox "合计成绩统计完成。"
End Sub