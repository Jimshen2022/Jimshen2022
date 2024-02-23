Sub aa()
    Dim arr, i&, n&, m&
    n = 0:m = 0
    arr = Sheets("数据源").[a1].CurrentRegion
    For i = 1 To UBound(arr)
        If arr(i, 1) = "∮25" Then
            If arr(i, 2) > 0.5 And arr(i, 2) < 0.7 Then n = n + 1
        End If
        If arr(i, 1) = "∮16" Then
            If arr(i, 2) > 0.6 And arr(i, 2) < 0.7 Then m = m + 1
        End If
    Next
    [a2] = n
    [b2] = m
End Sub





Sub aa2()
    Dim arr, i&, n&, m&, Nsum, Msum
    n = 0:m = 0
    arr = Sheets("数据源").[a1].CurrentRegion
    For i = 1 To UBound(arr)
        If arr(i, 1) = "∮25" Then
            If arr(i, 2) > 0.5 And arr(i, 2) < 0.7 Then
                n = n + 1 '计数
                Nsum = Nsum + arr(i, 2) '求和
            End If
        End If
        If arr(i, 1) = "∮16" Then
            If arr(i, 2) > 0.6 And arr(i, 2) < 0.7 Then
                m = m + 1
                Msum = Msum + arr(i, 2) '求和
            End If
        End If
    Next
    If n > 0 Then [a2] = Nsum / n '平均
    If m > 0 Then [b2] = Msum / m '平均
End Sub
