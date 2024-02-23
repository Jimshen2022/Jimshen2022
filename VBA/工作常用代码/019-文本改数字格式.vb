
Sub PullVarianceReport() '
    Dim wb As Workbook
    Dim arr, brr(), i&, j&, k&, nrow&, crr()
    
    't = Timer
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Sheet8.Activate
    Range("a2:ah10000").ClearContents
    
    Set wb = GetObject("C:\Users\jishen\Downloads\AT_MAPICS_HJ_KNQMAN.xlsx") '打开工作簿
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False
    
    ReDim brr(2 To UBound(arr), 1 To 13)
    For i = 2 To UBound(arr)
        For j = 1 To 13
            brr(i, j) = arr(i, j)
        Next
    Next
    With Sheet8
         .Range("aj1") = arr(1, 1)
         .Range("a1:b10000").NumberFormat = "@"
         .Range("a1").Resize(UBound(arr) - 1, 13) = brr
        '.Range("b2:b10000").Value = .Range("b2:b10000").Value
        'Columns("a:m").EntireColumn.AutoFit
         .Range("c2:m10000").NumberFormatLocal = ""
         .Range("c2:m10000").Value =  .Range("c2:m10000").Value
    End With
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    'MsgBox Format(Timer - t, "0.00" & "s")
    
End Sub