Sub Pull_AT_Orphaned_SN() '跨工作薄提取ORPHANED SN ---finished
    
    Dim wb As Workbook
    Dim arr, brr(), i&, j&, k&, nrow&, crr()
    
    't = Timer
    Application.ScreenUpdating = False
    'Sheet4.Range("a1").Value = "Data collected at:" & Format(Now(), "hhmm,mm-dd-yyyy")
    
    Sheet9.Cells.ClearContents
    
    Set wb = GetObject("C:\Users\jishen\Downloads\AT_ORPHANED.xlsx") '打开工作簿
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False
    ReDim brr(1 To UBound(arr), 1 To 15)
    For i = 1 To UBound(arr)
        For j = 1 To 15
            brr(i, j) = arr(i, j)
        Next
    Next
    
    Sheet9.Columns("a:h").NumberFormat = "@"
    Sheet9.Range("a1").Resize(UBound(arr)).Value = Application.Index(brr, , 1)
    Sheet9.Range("b1").Resize(UBound(arr)).Value = Application.Index(brr, , 2)
    Sheet9.Range("c1").Resize(UBound(arr)).Value = Application.Index(brr, , 3)
    Sheet9.Range("d1").Resize(UBound(arr)).Value = Application.Index(brr, , 4)
    Sheet9.Range("e1").Resize(UBound(arr)).Value = Application.Index(brr, , 5)
    Sheet9.Range("f1").Resize(UBound(arr)).Value = Application.Index(brr, , 7)
    Sheet9.Range("g1").Resize(UBound(arr)).Value = Application.Index(brr, , 8)
    Sheet9.Range("h1").Resize(UBound(arr)).Value = Application.Index(brr, , 10)
    
    Sheet9.Columns("a:o").EntireColumn.AutoFit
    
    Application.ScreenUpdating = True
    'MsgBox Format(Timer - t, "0.00" & "s")
    
End Sub