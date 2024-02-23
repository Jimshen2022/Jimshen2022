Sub Sitevlookup_A() ' Wanek3_DC_OUT column A for site
    
    
    't = Timer
    
    Dim d As Object, arr, brr, i&
    row = Sheet5.Range("i1048576").End(3).row
    
    Set d = CreateObject("scripting.dictionary")
    d.CompareMode = vbTextCompare '不区分字母大小写
    
    arr = Sheet8.Range("a1").CurrentRegion '数据源装入数组arr
    brr = Sheet5.Range("a1").CurrentRegion '查询区域数据装入数组brr
    
    For i = 1 To UBound(arr) '遍历数组arr
        d(arr(i, 1)) = arr(i, 2) '将item + up 作为key，装入字典
    Next
    
    For i = 2 To UBound(brr) '标题行不用查询，所以从第二行开始遍历查询数值brr
        If d.exists(brr(i, 5)) Then brr(i, 1) = d(brr(i, 5)) Else brr(i, 1) = "VIEW"
        '如果字典中存在item, '根据item 从字典中取UP值
        
        
    Next i
    
    'Sheet5.[a1].Resize(UBound(brr), UBound(brr, 2)) = brr   '(brr,2)是指数组的二纬下标
    With Sheet5.Range("a1").CurrentRegion
        
         .Value = brr
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlCenter
        
        'MsgBox "udpated~ " & Timer - t
        
    End With
    Set d = Nothing '释放字典
    
End Sub