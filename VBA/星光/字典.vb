Sub 前期绑定()  '需要在VBE\Reference\to add "Microsoft Scripting Runtime"
    Dim d As New Dictionary '声明一个字典对象
   
End Sub


Sub 后期绑定()
    Dim d As Object '声明一个对象
    Set d = CreateObject("scripting.dictionary") '创建对字典的引用
  
End Sub


# 比如——在使用前期绑定后，编写代码时系统会自动显示字典的成员列表，而后期绑定不会显示；
# 有些属性，前期绑定是支持的，但后期绑定不能使用（详情见文末说明）；另外，通常前期绑定的代码运算效率要比后期快一些。
# 但是——前期绑定的代码不适合发送给其它用户使用，毕竟其它用户未必会去手动绑定字典对象；因此后期绑定的方式兼容性更强些。
# 总结——建议编写代码时采用前期绑定，编写完成后，如需发送他人使用，再改为后期绑定，如此鱼和熊掌必可兼得矣~




#如何将数据装入字典

Sub 字典添加数据()
    Dim d As New Dictionary '声明一个字典对象
    d("看见星光") = 99
    d("Excel星球") = 98
End Sub


Sub 判断字典是否存在相同关键字()
    Dim d As New Dictionary '声明一个字典对象
    d("看见星光") = 99
    If Not d.Exists("看见星光") Then
        d("看见星光") = 59
    End If
End Sub


Sub 数据表数据存入字典()
    Dim d As New Dictionary
    Dim arr, i As Long
    arr = Worksheets("数据表").Range("a1").CurrentRegion
    For i = 2 To UBound(arr) '遍历数组元素
        d(arr(i, 1)) = arr(i, 2) '姓名是key，特长是item
    Next
End Sub


#如何将数据从字典取出
Sub 读取数据()
    Dim d As New Dictionary '声明一个字典对象
    d("看见星光") = 99
    d("Excel星球") = 98
    MsgBox d("Excel星球")
End Sub

# 需要根据A:B列的数据源，查询D列人名对应的特长，这就是所谓的条件查询了。


Sub 读取数据2()
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



# 如下图所示的数据表，A列人名存在重复值，需要去重复，获取不重复的人员名单，结果如D列。

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


#需要在D:E列，获取A:B列不重复的人名及其特长数据
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



#如何移除字典中的元素
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


# 如何遍历字典中的元素   如下图所示的数据表，需要筛选出人名重复出现次数大于2次的人员名单，以及相关出现次数，结果参考C:D列。

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


































































