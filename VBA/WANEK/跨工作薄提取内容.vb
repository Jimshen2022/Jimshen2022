Sub PullWanek3STO()  '跨工作薄提取内容
    
    Application.ScreenUpdating = False
    Dim wb As Workbook
    Dim arr
    
    Sheet3.Activate
    Columns("e:v").ClearContents
    
    Set wb = GetObject("C:\Users\jishen\Downloads\WANEK3STO.xlsx")      '打开工作簿
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False

    Sheet3.Columns("e:v").NumberFormat = "@"
    Sheet3.Range("e1").Resize(UBound(arr), 18) = arr

    Erase arr
    Application.ScreenUpdating = True
    
End Sub

Sub Pull_Wanek3_DC_361_OUT_trx()  '跨工作薄提取内容
    
    Application.ScreenUpdating = False
    
    Dim wb As Workbook
    Dim arr
    Sheet5.Activate
    Columns("i1:aj1048576").ClearContents
    
    Set wb = GetObject("C:\Users\jishen\Downloads\WANEK3DC361OUT.xlsx")      '打开工作簿
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False

    Sheet5.Columns("i:aj").NumberFormat = "@"
    Sheet5.Range("i1").Resize(UBound(arr), 28) = arr
    Erase arr
    Application.ScreenUpdating = True
    
End Sub

Sub Pull_Wanek3_111_in_trx()  '跨工作薄提取内容
    
    Application.ScreenUpdating = False
    Dim wb As Workbook
    Dim arr
    Sheet1.Activate
    Columns("e:af").ClearContents
    Set wb = GetObject("C:\Users\jishen\Downloads\WANEK3111IN.xlsx")      '打开工作簿
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False
    Sheet1.Columns("e:af").NumberFormat = "@"
    Sheet1.Range("e1").Resize(UBound(arr), 28) = arr

    Erase arr
    Application.ScreenUpdating = True

End Sub


Sub Pull_Wanek3_202_out_trx()  '跨工作薄提取内容
    
    Application.ScreenUpdating = False
    
    Dim wb As Workbook
    Dim arr, brr, i&, j&
    Sheet2.Activate
    Columns("e:az").ClearContents
    
    Set wb = GetObject("C:\Users\jishen\Downloads\WANEK3202OUT.xlsx")      '打开工作簿
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False
    ReDim brr(1 To UBound(arr), 1 To 28)
    For i = 1 To UBound(arr)
            If Not arr(i, 12) Like "C*" And Not arr(i, 12) Like "U*" And Not arr(i, 12) Like "M*" And Not arr(i, 12) Like "B1*" Then
                m = m + 1
                For j = 1 To 28
                    brr(m, j) = arr(i, j)
                Next
            End If
    Next
    Sheet2.Columns("e:af").NumberFormat = "@"
    Sheet2.Range("e1").Resize(UBound(arr), 28) = brr
    Erase arr
    Erase brr
    Application.ScreenUpdating = True

End Sub


Sub Pull_DC_202_in_trx()  '跨工作薄提取内容

    Application.ScreenUpdating = False
    Dim wb As Workbook
    Dim arr, brr, i&, j&

    Sheet4.Activate
    Columns("e:af").ClearContents
    
    Set wb = GetObject("C:\Users\jishen\Downloads\DC202IN.xlsx")      '打开工作簿
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False
    ReDim brr(1 To UBound(arr), 1 To 28)
    For i = 1 To UBound(arr)
            If arr(i, 9) Like "UL6*" Or arr(i, 9) Like "M*" Or arr(i, 9) = "To Location ID" Then
                m = m + 1
                For j = 1 To 28
                    brr(m, j) = arr(i, j)
                Next
            End If
    Next
    Sheet4.Columns("e:af").NumberFormat = "@"
    Sheet4.Range("e1").Resize(UBound(arr), 28) = brr
       
    Erase arr
    Erase brr
    Application.ScreenUpdating = True
    
End Sub

Sub Pull_BW_202_in_trx()  '跨工作薄提取内容

    Application.ScreenUpdating = False
    Dim wb As Workbook
    Dim arr, brr, i&, j&

    Sheet12.Activate
    Columns("e:af").ClearContents
    
    Set wb = GetObject("C:\Users\jishen\Downloads\DC202IN.xlsx")      '打开工作簿
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False
    ReDim brr(1 To UBound(arr), 1 To 28)
    For i = 1 To UBound(arr)
            If arr(i, 9) Like "UL9*" Or arr(i, 9) Like "B1*" Or arr(i, 9) = "To Location ID" Then
                m = m + 1
                For j = 1 To 28
                    brr(m, j) = arr(i, j)
                Next
            End If
    Next
    Sheet12.Columns("e:af").NumberFormat = "@"
    Sheet12.Range("e1").Resize(UBound(arr), 28) = brr
       
    Erase arr
    Erase brr
    Application.ScreenUpdating = True
    
End Sub


