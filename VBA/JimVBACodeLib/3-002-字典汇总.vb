
Sub itemclass_product() '跨工作薄提取itemclass and product ---finished
    Dim wb As Workbook
    Dim arr, brr(), i&, j&, k&, nrow&, crr(), d As Object, d2 As Object
    
    t = Timer
    Application.ScreenUpdating = False
    Sheet8.Activate
    Range("y2:z1048517").ClearContents
    nrow = Sheet8.Range("a1048576").End(3).row
    brr = Sheet8.Range("x2:ae" & nrow)
    
    Set d = CreateObject("scripting.dictionary")
    d.CompareMode = vbTextCompare '不区分字母大小写
    
    Set d2 = CreateObject("scripting.dictionary")
    d.CompareMode = vbTextCompare '不区分字母大小写
    
    Set wb = GetObject("D:\Document\02-Ashton\06-Tihi Calculate\AT_ITEM_2020.xlsb") '打开工作簿
    arr = wb.Worksheets("ITEM").[a1].CurrentRegion
    wb.Close False
    
    For i = 2 To UBound(arr) '遍历数组arr
        d(arr(i, 1)) = arr(i, 5) '将item + class 作为key，装入字典
        d2(arr(i, 1)) = arr(i, 27)
    Next
    
    For i = 1 To UBound(brr) '标题行不用查询，所以从第二行开始遍历查询数值brr
        If d.exists(brr(i, 1)) Then brr(i, 3) = d(brr(i, 1)) Else brr(i, 3) = "View"
        If d2.exists(brr(i, 1)) Then brr(i, 2) = d2(brr(i, 1)) Else brr(i, 2) = "View"
        
    Next
    Sheet8.Range("x2").Resize(UBound(brr), UBound(brr, 2)).Value = brr
    
    Set d = Nothing
    Set d2 = Nothing
    Set wb = Nothing
    
    
    Application.ScreenUpdating = True
    MsgBox Format(Timer - t, "0.00" & "s")
    
End Sub