Sub GenerateCCdate()
    Application.ScreenUpdating = False
    Dim i&, j&, x&, m%, n%, arr, ndate As Date
        
    With Sheet2
    
        'Inventory Cycle Count Interval--Finish cycle count for all locations in n days
        n = Sheet4.Range("b3").Value
        nrow = .Range("b1048576").End(3).Row
        x = Application.WorksheetFunction.RoundUp((nrow - 1) / n, 0)  'allocate locations in n days to finish
        .Range("c2:e1048576").ClearContents
        arr = .Range("a1:e" & nrow)
        
        'Schedule Starting Cycle count Date
        ndate = Sheet4.Range("b4").Value
        m = 1
    
            For i = 1 To UBound(arr) Step x
                For j = 1 To x
                    '如果 i+j 超过arr的边界则退出循环
                    If i + j > UBound(arr) Then
                      Exit For
                    Else
                         If arr(i + j, 2) <> "" Then
                            'cc sequence -- allocate all locations in x days
                            arr(i + j, 3) = m
                            
                             '如果是周6， CC Date 当前日期+2天
                             If Application.WorksheetFunction.Weekday(ndate) = 1 Then
                                ndate = ndate + 1
                                 arr(i + j, 4) = ndate
                             Else
                             
                                '其他星期， CC Date 当前日期+1天
                                arr(i + j, 4) = ndate
                             End If
                             
                             'WeekDay -- 检查schedule date是星期几
                             arr(i + j, 5) = Application.WorksheetFunction.Weekday(arr(i + j, 4)) - 1
        
                         Else
                             Exit For
                         End If
        
                     End If
        
                Next
                ndate = ndate + 1
                m = m + 1
            Next
            .Range("a1").Resize(UBound(arr), UBound(arr, 2)).Value = arr
            Erase arr
    End With
    Application.ScreenUpdating = True
'    MsgBox "finished!"
End Sub
