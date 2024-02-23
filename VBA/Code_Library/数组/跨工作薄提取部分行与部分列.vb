Sub Pull_AS_SN_HOLD()
    
    Dim wb As Workbook
    Dim arr, brr(), i&, j&, k&, nrow&, crr()
    
    't = Timer
    Application.ScreenUpdating = False
    'Sheet4.Range("a1").Value = "Data collected at:" & Format(Now(), "hhmm,mm-dd-yyyy")
    
    With Sheet5
         .Cells.Clear
        
        Set wb = GetObject("C:\Users\jishen\Downloads\AS_HOLD.xlsx") '打开工作簿
        arr = wb.ActiveSheet.[a1].CurrentRegion
        wb.Close False
        ReDim brr(1 To UBound(arr), 1 To UBound(arr, 2))
        For i = 1 To UBound(arr)
            If arr(i, 4) <> "Orphaned" Then
                k = k + 1
                For j = 1 To 15
                    brr(k, j) = arr(i, j)
                Next
            End If
        Next
        
         .Columns("a:i").NumberFormat = "@"
         .Range("a1").Resize(UBound(brr)).Value = Application.Index(brr, , 1)
         .Range("b1").Resize(UBound(brr)).Value = Application.Index(brr, , 2)
         .Range("c1").Resize(UBound(brr)).Value = Application.Index(brr, , 3)
         .Range("d1").Resize(UBound(brr)).Value = Application.Index(brr, , 4)
         .Range("e1").Resize(UBound(brr)).Value = Application.Index(brr, , 5)
         .Range("f1").Resize(UBound(brr)).Value = Application.Index(brr, , 7)
         .Range("g1").Resize(UBound(brr)).Value = Application.Index(brr, , 8)
         .Range("h1").Resize(UBound(brr)).Value = Application.Index(brr, , 10)
         .Range("i1").Resize(UBound(brr)).Value = Application.Index(brr, , 11)
         .Columns("a:i").EntireColumn.AutoFit
    End With
    Application.ScreenUpdating = True
    'MsgBox Format(Timer - t, "0.00" & "s")
    
End Sub
