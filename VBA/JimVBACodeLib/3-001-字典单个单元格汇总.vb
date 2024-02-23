Sub MetricsSummary()
    Application.ScreenUpdating = False
    ' 运行示例代码前需要引用“Microsoft Scripting Runtime”扩展库
    Dim i&, j&, k&, m&, n&, p&, x&, arr, d As Scripting.Dictionary
        
    nrow = Sheet3.Range("c1048576").End(3).Row
    Set d = CreateObject("scripting.dictionary")
    
    arr = Sheet3.Range("a1:l" & nrow)
    n = 0
    p = 0
    x = 0
    For i = 2 To nrow
        
        'by type to judge 0 or 1 to calculate No Variance Location
        If arr(i, 11) = "NoVariance" Then
            arr(i, 12) = 0
        Else
            arr(i, 12) = 1
        End If
        
        'calculate locations = 0 to get no variances location
        d(arr(i, 3)) = d(arr(i, 3)) + arr(i, 12)
        
        'summarized column "abs(variance)" on CC List Sheet
        n = n + arr(i, 10)
        
        'Variance column "Variance" on CC list Sheet
        p = p + arr(i, 9)
        
        'summarized "System Qty" on CC List
        x = x + arr(i, 5)
        
    Next
    
    '利用key,item 应该是成对的，也就是分别在 keys() items()的index是一样的。
    m = 0
    For j = LBound(d.items()) To UBound(d.items())
        If d.items(j) = 0 Then
              m = m + 1
        Else
        
        End If
    Next
    
    With Sheet9
    'No variance Location
        .Range("c2").Value = m
    
    'Total Locations
        .Range("c3").Value = d.Count
    
    'Absolute Variance Qty
        .Range("c4").Value = n

    'Variance Qty
        .Range("c5").Value = p
        
    '"Total Qty in system"
        .Range("c6").Value = x
        .Range("j1").Value = Sheet5.Range("m1").Value
    
    End With
    
    Erase arr
    Set d = Nothing
    ThisWorkbook.Save
    Application.ScreenUpdating = True


End Sub

