Sub Splits()

Application.ScreenUpdating = False
Dim i&, j%, arr, srr, crr()

Sheet7.Activate
arr = Sheet7.Range("a1").CurrentRegion
ReDim crr(1 To UBound(arr), 1 To 7)

For i = 1 To UBound(arr)
        srr = Split(arr(i, 1), "/")
    For j = 0 To UBound(srr)     ' 这里的srr为一维数组
        crr(i, j + 1) = srr(j)   ' 遍历srr, 从0开始
    Next
           
Next
Sheet7.Range("i1").Resize(UBound(arr), 7).Value = crr
Application.ScreenUpdating = True

End Sub