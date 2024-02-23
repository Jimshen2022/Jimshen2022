
'数组装入字典，再遍历字典的items与key，再SPLIT存放于数组
'D:\Document\01-Wanvog\10-CC\Anual CC\2021\HJ vs Mapics Comparing - 2021.xlsx

Sub DistinctSN1()
    
    Application.ScreenUpdating = False
    t = Timer
    Dim aKey, aItem, kRes, iRes, arr, i&, j&, d As Object

    Sheet3.Activate
    Sheet3.Cells.Clear
    arr = Sheet2.Range("a1").CurrentRegion
    
    Set d = CreateObject("scripting.dictionary")
    d.CompareMode = vbTextCompare '不区分字母大小写
    
    For i = 2 To UBound(arr) '遍历数组arr
        If Not arr(i, 9) Like "RP0%" Then
            d(arr(i, 2)) = arr(i, 1) & ";" & arr(i, 3) & ";" & arr(i, 5) & ";" & arr(i, 6) & ";" & arr(i, 8) & ";" & arr(i, 9) & ";" & arr(i, 11) '将SN 作为key，装入字典
        End If
    Next
  
    '遍历字典的items,存放于aRes数组
    aItem = d.items
    aKey = d.Keys
    
    ReDim iRes(1 To d.Count + 1, 1 To 7)  '结果数组
    ReDim kRes(1 To d.Count + 1, 1 To 1)  '结果数组
    
    
    For i = 0 To UBound(aKey)
        For j = 1 To 1
           kRes(i + 1, j) = aKey(i)
        Next
    Next
       
    With Sheet3
        .Columns("a:g").NumberFormat = "@"
        .Columns("a:h").AutoFit
        .Range("a2").Resize(UBound(iRes), 1) = kRes
    End With
    
    
    For i = 0 To UBound(aItem)
        For j = 0 To 6
           iRes(i + 1, j + 1) = Split(aItem(i), ";")(j)
        Next
    Next
    
    With Sheet3
        .Columns("a:g").NumberFormat = "@"
        .Columns("a:h").AutoFit
        .Range("b2").Resize(UBound(iRes), 7) = iRes
    End With
    
    Set d = Nothing
    Erase aItem
    Erase aKey
    Erase kRes
    Erase iRes
    Erase arr
    ThisWorkbook.Save
    Application.ScreenUpdating = True
    MsgBox Format(Timer - t, "0.00" & "s")

    
End Sub