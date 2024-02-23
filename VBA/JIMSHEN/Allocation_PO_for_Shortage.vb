Sub Qiliao() 'ready 12
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = " Allocation by PO and MO, please wait ......"
    
    
    Dim arr1(), arr2()
    Dim a%, b%
    
    Sheets("8.Shortage trips").Select
    'Range("b1", Cells(Rows.Count, "g").End(xlUp)).Sort [d1], xlAscending, [b1], , xlAscending, , , xlYes  '按款号与日期排序
    arr1 = Range("a2", Cells(Cells(Rows.Count, "g").End(xlUp).Row, "h")) '数组赋值
    Sheets("9.Supply").Select
    'Range("a1", Cells(Rows.Count, "c").End(xlUp)).Sort [a1], xlAscending, [c1], , xlAscending, , , xlYes   '按款号与日期排序
    arr2 = Range("a2", Cells(Rows.Count, "c").End(xlUp)) '数组赋值
    
    'For a = LBound(arr1) To UBound(arr1)
    '    arr1(a, 6) = arr1(a, 6) * 1   '将文字改为数值
    'Next
    
    Sheets("8.Shortage trips").Select
    Range("h2:h66365").ClearContents
    
100
    For a = LBound(arr1) To UBound(arr1)
        For b = LBound(arr2) To UBound(arr2)
            If arr1(a, 4) = arr2(b, 1) And arr1(a, 6) <= arr2(b, 2) And arr1(a, 6) <> 0 Then '订单款号与PO款号相同，且 订单需求<=PO数量 且 订单需求不为0
                arr2(b, 2) = arr2(b, 2) - arr1(a, 6) 'PO数量 = PO数量 - 订单需求
                arr1(a, 6) = arr1(a, 6) - arr1(a, 6) '订单需求  = 订单需求 - 订单需求
                Sheets("8.Shortage trips").Cells(a + 1, "h") = arr2(b, 3) '预计齐料日 = PO的预计进料日期   ‘cells(a+1)是为了跟数组第一行对齐
                arr1(a, 8) = arr2(b, 3)
            ElseIf arr1(a, 4) = arr2(b, 1) And arr1(a, 6) > arr2(b, 2) And arr2(b, 2) <> 0 Then '亦或，订单款号与PO款号相同，且 订单需求>PO数量 且 订单需求不为0
                arr1(a, 6) = arr1(a, 6) - arr2(b, 2) '订单缺货数量 = 订单缺货数量 - PO数量
                arr2(b, 2) = arr2(b, 2) - arr2(b, 2) 'PO数量 = PO数量 - PO数量
                If arr1(a, 6) <> 0 Then '如果订单需求<>0, 回到100
                    GoTo 100
                End If
            End If
        Next
    Next
    
    
    
    Sheets("8.Shortage trips").Select
    
    Dim i%
    For i = 2 To [h66365].End(3).Row
        If Cells(i, "h") = "" Then Cells(i, "h") = #12 / 31 / 2099#
    Next
    
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    
End Sub