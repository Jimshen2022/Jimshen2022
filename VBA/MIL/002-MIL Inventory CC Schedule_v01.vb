Sub GenerateCCdate()
    Application.ScreenUpdating = False
    Dim i&, j&, x&, m%, n%, arr, ndate As Date
        
    With Sheet2
    
        'Inventory Cycle Count Interval--Finish cycle count for all locations in n days
        n = Sheet4.Range("b3").Value
        nrow = .Range("a1048576").End(3).Row
        x = Application.WorksheetFunction.RoundUp((nrow - 1) / n, 0)  'allocate locations in n days to finish
        .Range("b2:d" & nrow).ClearContents
        arr = .Range("a1:d" & nrow)
        
        'Schedule Starting Cycle count Date
        ndate = Sheet4.Range("b4").Value
        m = 1
    
            For i = 1 To UBound(arr) Step x
                For j = 1 To x
                    '如果 i+j 超过arr的边界则退出循环
                    If i + j > UBound(arr) Then
                      Exit For
                    Else
                         If arr(i + j, 1) <> "" Then
                            'cc sequence -- allocate all locations in 22 days
                            arr(i + j, 2) = m
                            
                             '如果是周6， CC Date 当前日期+2天
                             If Application.WorksheetFunction.Weekday(ndate) = 7 Then
                                ndate = ndate + 2
                                 arr(i + j, 3) = ndate
                             Else
                             
                                '其他星期， CC Date 当前日期+1天
                                arr(i + j, 3) = ndate
                             End If
                             
                             'WeekDay -- 检查schedule date是星期几
                             arr(i + j, 4) = Application.WorksheetFunction.Weekday(arr(i + j, 3)) - 1
        
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
    ThisWorkbook.Save
    Application.ScreenUpdating = True
    MsgBox "finished!"
End Sub
