
Sub Dicttl2()
    Dim d As Object, arr, brr, crr, i&, j, k&
    Set d = CreateObject("scripting.dictionary")
    '后期字典
    'd.CompareMode = vbTextCompare
    '不区分字母大小写
    arr = Range("a1:b" & Cells(Rows.Count, 1).End(xlUp).Row)
    '数据源装入数组arr
    ReDim crr(1 To UBound(arr), 1 To 3)
    '声明数组crr放置数据统计结果。1列姓名2列次数3列总成绩。姓名列可以省略。
    For i = 1 To UBound(arr)
    '先遍历数据源arr
        If Not d.exists(arr(i, 1)) Then
        '如果字典中不存在姓名……
            k = k + 1
            '累加不重复人名个数，可以先理解成人名在数组crr中的序列号
            d(arr(i, 1)) = k
            '将数组crr中的序列位置作为item装入字典，以便以后根据人名读取处理
            crr(k, 1) = arr(i, 1) '姓名
            crr(k, 2) = 1 '考试次数
            crr(k, 3) = Val(arr(i, 2))
            '考试成绩。val函数提取纯数值，如果是纯文本值则计算为0，该函数可以避免文本值数学运算时出错。
        Else
        '如果字典中存在相关人名
            j = d(arr(i, 1)) '读取人名在数组crr中的序列号
            crr(j, 2) = crr(j, 2) + 1
            '原次数+1
            crr(j, 3) = crr(j, 3) + Val(arr(i, 2))
            '累加成绩
        End If
    Next
    brr = Range("d1:f" & Cells(Rows.Count, 4).End(xlUp).Row)
    '查询区域装入数组brr
    For i = 2 To UBound(brr)
        If d.exists(brr(i, 1)) Then
        '如果字典中存在查询的姓名
            j = d(brr(i, 1))
            '姓名在数组brr中的序列号
            brr(i, 2) = crr(j, 2) '考试次数
            brr(i, 3) = crr(j, 3) '总成绩
        Else
        '否则返回空文本
            brr(i, 2) = ""
            brr(i, 3) = ""
        End If
    Next
    With Range("d1:f" & Cells(Rows.Count, 4).End(xlUp).Row)
        .NumberFormat = "@"
        '设置文本格式，防止某些文本数值数据变形
        .Value = brr
        'brr数组放回单元格区域
    End With
    Set d = Nothing
    '释放字典
    MsgBox "数据统计完成。"
End Sub