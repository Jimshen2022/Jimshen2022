Sub PullATtrx() '跨工作薄提取内容 ---finished
    Dim wb As Workbook
    Dim arr, brr(), i&, j&, k&, nRow&, crr()
    
    t = Timer
    Application.ScreenUpdating = False
    'Sheet2.Range("a1").Value = "Data collected at:" & Format(Now(), "hhmm,mm-dd-yyyy")
    
    Sheet2.Activate
    Range("a1:ab1048517").ClearContents
    
    Set wb = GetObject("C:\Users\jishen\Downloads\AT_TRX.xlsx") '打开工作簿
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False
    ReDim brr(1 To UBound(arr), 1 To 28)
    For i = 1 To UBound(arr)
        For j = 1 To 28
            brr(i, j) = arr(i, j)
        Next
    Next
    Sheet2.Range("a1").Resize(UBound(arr), 28) = brr
    
    Sheet2.Activate
    Columns("a:ab").NumberFormat = "@"
    Columns("a:ab").EntireColumn.AutoFit
    Sheet2.Select
    
    Application.ScreenUpdating = True
    MsgBox Format(Timer - t, "0.00" & "s")
    
End Sub

Sub delet_AT_Trx() '数组删除行
    
    Dim i&, j&, nRow&, m&, arr(), brr()
    With Sheet2
        nRow =  .Range("a1048576").End(xlUp).Row
        arr =  .Range("a2:ab" & nRow).Value
        ReDim brr(1 To nRow - 1, 1 To 28)
        For i = 2 To nRow - 1
            If arr(i, 16) = "252" Or arr(i, 16) = "364" Or arr(i, 16) = "202" Or arr(i, 16) = "321" Or arr(i, 16) = "304" Or arr(i, 16) = "254" Or arr(i, 16) = "372" Or arr(i, 16) = "856" Or arr(i, 16) = "312" Or arr(i, 16) = "152" Or arr(i, 16) = "262" Then
                m = m + 1
                For j = 1 To 28
                    brr(m, j) = arr(i, j)
                Next
            End If
        Next
         .Range("a2:ab" & nRow).Value = brr
    End With
End Sub
