Sub 预计齐料()
    
    Application.ScreenUpdating = False
    
    ends = Range("B65536").End(xlUp).Row
    
    Set s = Sheets("采购订单")
    
    ends2 = s.Range("B65536").End(xlUp).Row
    Sheet2.Range("e1") = "辅助列"
    Sheet2.Range("e2:e66563").ClearContents
    For i = 2 To ends
        n = 0
        For j = 2 To ends2
            
            If Range("B" & i) = s.Range("B" & j) And s.Range("C" & j) <> s.Range("E" & j) Then '如果销售订单的款号=采购订单的款号 and 采购订单的数量 <> 采购订单的E栏
                
                If s.Range("E" & j) <> "" Then '如果采购订单的E栏不为空白
                    
                    x = s.Range("E" & j) 'X = E栏(辅助栏）
                Else
                    
                    x = s.Range("C" & j) '否则 X = 采购订单数量
                End If
                
                If x - Range("D" & i) + n >= 0 Then '如果x-订单缺货数量+n >=0
                    
                    s.Range("E" & j) = x - Range("D" & i) + n '采购订单E栏 = x - 销售订单数量 + n
                    Range("E" & i) = s.Range("A" & j) '预计齐料日期 = 采购订单的交期
                    Exit For '退出循环
                    
                Else '如果x-订单缺货数量+n <0
                    
                    s.Range("E" & j) = s.Range("C" & j) '采购订单E栏 = 采购订单数量
                    n = n + x '并且累加采购订单数量
                    
                End If
                
            End If
            
        Next j
        
    Next i
    
    For k = 2 To ends
        
        If Range("E" & k) = "" Then
            
            Range("E" & k) = "PO数量不足"
            
        End If
        
    Next k
    
    MsgBox "更新完成"
    
    
End Sub



Sub 预计齐料2()
    
    Application.ScreenUpdating = False
    Dim i%, k%, n%, x%, arr(), brr()
    ends = Range("B65536").End(xlUp).Row
    arr = Sheet1.Range("a1").CurrentRegion 'Shortage
    brr = Sheet2.Range("a1").CurrentRegion 'Supply
    Sheet2.Range("e1") = "辅助列"
    Sheet2.Range("e2:e66563").ClearContents
    
    For i = 2 To UBound(arr)
        n = 0
        For j = 2 To UBound(brr)
            
            If arr(i, 2) = brr(j, 2) And brr(j, 3) <> brr(j, 5) Then '如果销售订单的款号=采购订单的款号 and 采购订单的数量 <> 采购订单的E栏
                If brr(j, 5) <> "" Then '如果采购订单的E栏不为空白
                    x = brr(j, 5) 'X = E栏(辅助栏）
                Else
                    x = brr(j, 3) '否则 X = 采购订单数量
                End If
                
                
                If x - arr(i, 4) + n >= 0 Then '如果x-订单缺货数量+n >=0
                    brr(j, 5) = x - arr(i, 4) + n '采购订单E栏 = x - 销售订单数量 + n
                    arr(i, 5) = brr(j, 1) '预计齐料日期 = 采购订单的交期
                    Exit For '退出循环
                    
                Else '如果x-订单缺货数量+n <0
                    brr(j, 5) = brr(j, 3) '采购订单E栏 = 采购订单数量
                    n = n + x '并且累加采购订单数量
                End If
            End If
        Next j
    Next i
    
    Sheet1.Columns("a:d").NumberFormat = "@"
    Sheet1.Range("a1").Resize(UBound(arr), UBound(arr, 2)).Value = arr
    
    For k = 2 To ends
        If Sheet1.Range("E" & k) = "" Then
            Sheet1.Range("E" & k) = "PO数量不足"
        End If
    Next k
    
    MsgBox "更新完成"
    
    
End Sub