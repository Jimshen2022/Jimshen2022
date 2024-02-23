Sub PullWanek1STO() '¿ç¹¤×÷±¡ÌáÈ¡ÄÚÈÝ
    Dim wb As Workbook
    Dim arr, brr(), i&, j&, k&, nRow&, crr()
    
    't = Timer
    'Application.ScreenUpdating = False
    Sheet3.Activate
    Columns("e:v").ClearContents
    
    Set wb = GetObject("C:\Users\jishen\Downloads\WANEK1STO.xlsx") '´ò¿ª¹¤×÷²¾
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False
    ReDim brr(1 To UBound(arr), 1 To 18)
    For i = 1 To UBound(arr)
        If Mid(arr(i, 2), 1, 1) <> "B" Then
            m = m + 1
            For j = 1 To UBound(arr, 2)
                brr(m, j) = arr(i, j)
            Next
        End If
    Next
    
    
    Sheet3.Range("e1").Resize(UBound(arr), 18) = brr
    
    Sheet3.Activate
    Columns("e:v").NumberFormat = "@"
    Columns("a:v").EntireColumn.AutoFit
    Sheet3.Select
    
    Erase arr
    Erase brr
    
End Sub



Sub 数组条件筛选()
    arr = [a1].CurrentRegion
    ReDim brr(1 To UBound(arr), 1 To UBound(arr, 2))
    For i = 2 To UBound(arr)
        If arr(i, 3) = "左" And arr(i, 4) = "南方公司" Then
            m = m + 1
            For j = 1 To UBound(arr, 2)
                brr(m, j) = arr(i, j)
            Next
        End If
    Next
    If m Then [f1].Resize(UBound(brr), UBound(brr, 2)) = brr
End Sub



