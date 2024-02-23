Sub STOtype1delete()  '数组删除行 --删除STO type1栏的非storage行 --- finished
    Application.ScreenUpdating = False
    Dim i&, j&, nrow&, m&, arr(), brr()
    With Sheet2
        nrow = .Range("p1048576").End(xlUp).Row
        arr = .Range("o2:af" & nrow).Value
        ReDim brr(1 To nrow, 1 To 18)
        For i = 1 To nrow - 1
            If arr(i, 15) = "STORAGE" And arr(i, 7) <> "SP001AA1" And arr(i, 7) <> "QA001VD1" And arr(i, 7) <> "NG001AD1" Then
                m = m + 1
                For j = 1 To 18
                    brr(m, j) = arr(i, j)
                Next
            End If
        Next
        .Range("o2:af" & nrow).Value = brr
    End With
    Application.ScreenUpdating = True
End Sub