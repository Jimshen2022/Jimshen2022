
'在VBE编辑器的[工具] 选项卡下 ，单击[引用] ，在打开的[引用] 对话框中勾选 "Microsoft Scripting Runtim" 选项 ，单击[确定] 按钮 ，关闭对话框即可 。
Sub 前期绑定()
    Dim d As New Dictionary '声明一个字典对象
    ……
End Sub


Sub 后期绑定()
    Dim d As Object '声明一个对象
    Set d = CreateObject("scripting.dictionary") '创建对字典的引用
    ……
End Sub

Sub 字典添加数据()
    Dim d As New Dictionary '声明一个字典对象
    d("看见星光") = 99
    d("Excel星球") = 98
End Sub



Sub 数据表数据存入字典()
    Dim d As New Dictionary
    Dim arr, i As Long
    arr = Worksheets("数据表").Range("a1").CurrentRegion
    For i = 2 To UBound(arr) '遍历数组元素
        d(arr(i, 1)) = arr(i, 2) '姓名是key，特长是item
    Next
End Sub


Sub 读取数据()
    Dim d As New Dictionary '声明一个字典对象
    d("看见星光") = 99
    d("Excel星球") = 98
    MsgBox d("Excel星球")
End Sub


Sub 判断字典是否存在相同关键字()
    Dim d As New Dictionary '声明一个字典对象
    d("看见星光") = 99
    If Not d.Exists("看见星光") Then
        d("看见星光") = 59
    End If
End Sub




Sub 条件查询()
    Dim d As New Dictionary
    Dim arr, brr, i As Long
    arr = Range("a1").CurrentRegion '数据源
    For i = 2 To UBound(arr) '遍历数组，数据装入字典
        d(arr(i, 1)) = arr(i, 2) 'key是人名,item是特长
    Next
    brr = Range("d1:e" & Cells(Rows.Count, "d").End(xlUp).Row) '查询区域
    For i = 2 To UBound(brr) '遍历查询值
        If d.Exists(brr(i, 1)) Then '如果字典存在查询值
            brr(i, 2) = d(brr(i, 1)) '获取人名对应的条目
        Else
            brr(i, 2) = "查无此人"
        End If
    Next
    Range("d1:e" & Cells(Rows.Count, "d").End(xlUp).Row) = brr
    Set d = Nothing
End Sub



Sub 去重复()
    Dim d As New Dictionary
    Dim arr, i As Long
    arr = Range("a1").CurrentRegion
    For i = 1 To UBound(arr)
        If Not d.Exists(arr(i, 1)) Then
            d(arr(i, 1)) = ""
        End If
    Next
    Range("d:d").ClearContents
    Range("d1").Resize(d.Count, 1) = Application.Transpose(d.Keys)
    Set d = Nothing
End Sub


Sub 去重复2()
    Dim d As New Dictionary
    Dim arr, i As Long
    arr = Range("a1").CurrentRegion
    For i = 1 To UBound(arr)
        If Not d.Exists(arr(i, 1)) Then
            d(arr(i, 1)) = arr(i, 2)
        End If
    Next
    Range("d:e").ClearContents
    Range("d1").Resize(d.Count, 2) = Application.Transpose(Array(d.Keys, d.Items))
    Set d = Nothing
End Sub


Sub 移除指定关键字()
    Dim d As New Dictionary '声明一个字典对象
    d("看见星光") = 99
    d("Excel星球") = 98
    d.Remove "看见星光"
End Sub



Sub 全部移除字典中的元素()
    Dim d As New Dictionary
    Dim arr, i As Long, j As Long
    arr = Range("a1").CurrentRegion
    For i = 2 To UBound(arr) '遍历行
        d.RemoveAll '移除字典中所有的元素
        For j = 2 To UBound(arr, 2) - 1 '遍历列
            If Not d.Exists(arr(i, j)) Then
                d(arr(i, j)) = "" '将不重复的号码存入字典
            End If
        Next
        arr(i, UBound(arr, 2)) = VBA.Join(d.Keys, ",") '合并为一个字符串
    Next
    Range("a1").CurrentRegion = arr
    Set d = Nothing
End Sub


Sub 遍历字典元素_索引法()
    Dim d As New Dictionary
    Dim arr, aKey, aRes, i As Long, k As Long
    arr = Range("a1").CurrentRegion
    For i = 2 To UBound(arr)
        d(arr(i, 1)) = d(arr(i, 1)) + 1
    Next
    aKey = d.Keys
    ReDim aRes(1 To d.Count, 1 To 2) '结果数组
    For i = 0 To UBound(aKey)
        If d(aKey(i)) > 2 Then '次数大于2次
            k = k + 1
            aRes(k, 1) = aKey(i)
            aRes(k, 2) = d(aKey(i))
        End If
    Next
    Range("c:c").ClearContents
    Range("c1") = "重复2次以上的人名"
    Range("c2").Resize(k, 2) = aRes
    Set d = Nothing
End Sub


Sub 不区分字母大小写()
    Dim d As New Dictionary
    d.CompareMode = TextCompare
    d("a") = 1
    MsgBox d("A")
End Sub


Sub 不区分字母大小写2()
    Dim d As New Dictionary
    d(LCase("a")) = 1 'LCase将字母统一转换为小写
    MsgBox d(LCase("A"))    
End Sub


Sub 数据类型2()
    Dim d As New Dictionary
    Dim strKey As String '定义一个字符串类型的变量
    strKey = "1"
    d(strKey) = "爱就一个字"
    strKey = "2"
    d(strKey) = "我要说两次"
    strKey = 1 '强制转换为字符串
    MsgBox d.Exists(strKey) '结果返回True
End Sub


