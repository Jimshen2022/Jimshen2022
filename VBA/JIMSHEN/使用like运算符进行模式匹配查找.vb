Sub RngLike()
    ' 5-2 使用like运算符进行模式匹配查找
    
    Dim rng As Range, r&
    r = 1
    Sheet1.Range("a:a").ClearContents
    For Each rng In Sheet2.Range("a1:a10000")
        If rng.Text Like "*a*" Then
            Sheet1.Cells(r, 1) = rng.Text
            r = r + 1
        End If
    Next
    Set rng = Nothing
End Sub
