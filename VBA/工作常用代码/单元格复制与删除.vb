Sub closedtofinished() '根据条件将单元格整行从一个sheet复制到另一个sheet, 并将原单元格整行删除
    
    Dim i%, m%
    
    For i = 2 To Sheet2.Range("a66563").End(3).Row
        If Sheet2.Range("k" & i) = "Closed" Then
            k = Sheet3.Range("a66563").End(3).Row
            Sheet2.Range("k" & i).EntireRow.Copy Sheet3.Range("a" & k + 1)
        End If
        
    Next
    
    For m = Sheet2.Range("a66563").End(3).Row To 2 Step -1
        If Sheet2.Range("k" & m) = "Closed" Then
            Sheet2.Range("k" & m).EntireRow.Delete
        End If
        
    Next
    
End Sub
