Sub delet_AT_Trx() '数组删除行
    
    Dim i&, j&, nRow&, m&, arr(), brr()
    With Sheet7
        nRow =  .Range("a1048576").End(xlUp).Row
        arr =  .Range("a2:ab" & nRow).Value
        ReDim brr(1 To nRow - 1, 1 To 28)
        For i = 2 To nRow - 1
            If arr(i, 3) <> "UPH" And arr(i, 7) Like "VR*" And (arr(i, 16) = "856" Or arr(i, 16) = "152" Or arr(i, 16) = "202" Or arr(i, 16) = "252" Or arr(i, 16) = "254" Or arr(i, 16) = "262" _
                         Or ((arr(i, 16) = "304" Or arr(i, 16) = "312" Or arr(i, 16) = "321" Or arr(i, 16) = "364" Or arr(i, 16) = "372") And arr(i, 15) Like "000000*")) Then
                
                m = m + 1
                For j = 1 To 28
                    brr(m, j) = arr(i, j)
                Next
            End If
        Next
        Stop
         .Range("a2:ab" & nRow).Value = brr
    End With
    
    MsgBox "UPDATED"
End Sub