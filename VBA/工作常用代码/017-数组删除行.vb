Sub Wanek3202outdelete() '数组删除行
    
    Dim i&, j&, nRow&, m&, arr(), brr()
    With Sheet2
        nRow =  .Range("p1048576").End(xlUp).row
        arr =  .Range("a2:af" & nRow).Value
        ReDim brr(1 To nRow, 1 To 32)
        For i = 1 To nRow - 1
            If Not arr(i, 16) Like "CN*" And Not arr(i, 16) Like "UL*" And Not arr(i, 16) Like "M*" Then
                m = m + 1
                For j = 1 To 32
                    brr(m, j) = arr(i, j)
                Next
            End If
        Next
         .Range("a2:af" & nRow).Value = brr
    End With
End Sub