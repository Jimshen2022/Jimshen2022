'Pull_HJ VS Mapics_Variance_REPORT, coded by JimShen on Sep.04.2021

Sub Pull_HJ_vs_Mapics()
    Dim wb As Workbook
    Dim arr, brr(), i&, j&, k&, nrow&, crr()
        
    t = Timer
    Application.ScreenUpdating = False
    
    Sheet21.Select
    Cells.Clear
    
    
    Set wb = GetObject("C:\Users\jishen\Downloads\Mapics_vs.xlsx") '?????
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False
    ReDim brr(2 To UBound(arr), 1 To UBound(arr, 2))

    For i = 2 To UBound(arr)
'            brr(i, 1) = arr(i, 2)
'            brr(i, 2) = arr(i, 3)
'            brr(i, 3) = arr(i, 7)
        For j = 1 To UBound(arr, 2)
            brr(i, j) = arr(i, j)
        Next
    Next
    
    Columns("a:b").NumberFormat = "@"
    'Columns("c:d").NumberFormat = "@"
    Sheet21.Range("a1").Resize(UBound(arr) - 1, UBound(arr, 2)) = brr
    Range("z1").Value = arr(1, 1)
    Range("c:d").EntireColumn.Insert
    Range("c1:d1").Value = Array("Reason", "ABS")
    
    crr = Range("a1").CurrentRegion
    For k = 2 To UBound(crr)
        crr(k, 4) = Abs(crr(k, 5))
    Next
    Range("a1").Resize(UBound(crr), UBound(crr, 2)).Value = crr
    
    Range("d2:m" & UBound(crr)).Value = Range("d2:m" & UBound(crr)).Value
    Columns("a:ab").EntireColumn.AutoFit
    
    Application.ScreenUpdating = True
    MsgBox Format(Timer - t, "0.00" & "s")
    
End Sub


