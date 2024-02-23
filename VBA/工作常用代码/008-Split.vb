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






'原文链接：https://blog.csdn.net/taller_2000/article/details/86713631

'请注意变量声明语句，用于保存结果的数组，可以使用如下两种方式：Variant变量或者字符串数组，但是不可以声明为Variant数组。
Sub Demo1()
    Dim strString As String
    Dim varResult As Variant
    Dim arrResult() As String
    strString = "Good good study day day up"
    varResult = VBA.Split(strString)
    arrResult = VBA.Split(strString)
End Sub


'一般情况下，都无须指定LIMIT参数，下面看一个使用LIMIT参数的例子。对于一些国外地址如：888, Ocean Wind Rd, Markham, V4A，需要拆分为888，Ocean Wind 'Rd和Markham, V4A，而不是拆分为4段，此时就需要设置LIMIT参数为3。
Sub Demo2()
    Dim strString As String
    Dim arrResult() As String
    Dim i As Integer
    strString = "888, Ocean Wind Rd, Markham, V4A"
    arrResult = VBA.Split(strString, ",", 3)
    For i = LBound(arrResult) To UBound(arrResult)
        Debug.Print Trim(arrResult(i))
    Next i
End Sub



Sub Demo3()
    Dim strString As String
    Dim arrResult() As String
    Dim i As Integer
    strString = "AAAsBBBSCCCsDDD"
    arrResult = VBA.Split(strString, delimiter:="s", compare:=vbTextCompare)
    For i = LBound(arrResult) To UBound(arrResult)
        Debug.Print Trim(arrResult(i))
    Next i
End Sub



Sub Demo4()
    Dim strString As String
    Dim arrResult() As String
    Dim i As Integer
    strString = "AAAsBBBSCCCsDDD"
    arrResult = VBA.Split(strString, delimiter:="s", compare:=vbBinaryCompare)
    For i = LBound(arrResult) To UBound(arrResult)
        Debug.Print Trim(arrResult(i))
    Next i
End Sub