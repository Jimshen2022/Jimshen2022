Sub Pull_AT_SN_InWarehouse() '跨工作薄提取NG_SN_LIST --- finished
    Dim wb As Workbook
    Dim Arr, brr(), i&, j&, k&, nrow&, crr()
    
    't = Timer
    Application.ScreenUpdating = False
    'Sheet4.Range("a1").Value = "Data collected at:" & Format(Now(), "hhmm,mm-dd-yyyy")
    
    Sheet22.Cells.ClearContents
    Set wb = GetObject("C:\Users\jishen\Downloads\AT_SN.xlsx") '打开工作簿
    Arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False
    ReDim brr(1 To UBound(Arr), 1 To 15)
    For i = 1 To UBound(Arr)
        For j = 1 To 15
            brr(i, j) = Arr(i, j)
        Next
    Next
    Sheet22.Range("a1").Resize(UBound(Arr)).Value = Application.Index(brr, , 1)
    Sheet22.Range("b1").Resize(UBound(Arr)).Value = Application.Index(brr, , 2)
    Sheet22.Range("c1").Resize(UBound(Arr)).Value = Application.Index(brr, , 3)
    Sheet22.Range("d1").Resize(UBound(Arr)).Value = Application.Index(brr, , 4)
    Sheet22.Range("e1").Resize(UBound(Arr)).Value = Application.Index(brr, , 5)
    Sheet22.Range("f1").Resize(UBound(Arr)).Value = Application.Index(brr, , 8)
    Sheet22.Range("g1").Resize(UBound(Arr)).Value = Application.Index(brr, , 10)
    '论坛上有一个帖子，题目大概是,数组的用法提到
    '
    '提取数组行:
    '
    'Dim Arr(1 To 6, 1 To 18)
    '
    'Arr1 = Application.Index(Arr, 6) '提取第六行数据
    '
    '提取数组列
    '
    'Arr1 = Application.Index(Arr, 0, 15) '提取第15列数据
    '
    '数据量大时，用这个方法速度会比循环快
    
    
    Sheet22.Activate
    Columns("a:o").NumberFormat = "@"
    Columns("a:o").EntireColumn.AutoFit
    
    Application.ScreenUpdating = True
    'MsgBox Format(Timer - t, "0.00" & "s")
    
End Sub

Sub Pull_AT_SN_Loaded() '跨工作薄提取NG_SN_LOADED_LIST --- finished
    
    Application.ScreenUpdating = False
    Dim wb As Workbook
    Dim Arr, brr(), i&, j&, k&, nrow&, crr()
    nrow = Sheet22.Range("a1048576").End(3).row
    
    Set wb = GetObject("C:\Users\jishen\Downloads\AT_SN-2.xlsx") '打开工作簿
    Arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False
    ReDim brr(2 To UBound(Arr), 1 To 15)
    For i = 2 To UBound(Arr)
        For j = 1 To 15
            brr(i, j) = Arr(i, j)
        Next
    Next
    Sheet22.Range("a1048576").End(3).Offset(1).Resize(UBound(Arr) - 1).Value = Application.Index(brr, , 1)
    Sheet22.Range("b1048576").End(3).Offset(1).Resize(UBound(Arr) - 1).Value = Application.Index(brr, , 2)
    Sheet22.Range("c1048576").End(3).Offset(1).Resize(UBound(Arr) - 1).Value = Application.Index(brr, , 3)
    Sheet22.Range("d1048576").End(3).Offset(1).Resize(UBound(Arr) - 1).Value = Application.Index(brr, , 4)
    Sheet22.Range("e1048576").End(3).Offset(1).Resize(UBound(Arr) - 1).Value = Application.Index(brr, , 5)
    Sheet22.Range("f1048576").End(3).Offset(1).Resize(UBound(Arr) - 1).Value = Application.Index(brr, , 8)
    Sheet22.Range("g1048576").End(3).Offset(1).Resize(UBound(Arr) - 1).Value = Application.Index(brr, , 10)
    
    'Sheet22.Range("a" & nrow).Offset(1).Resize(UBound(Arr) - 1, 15) = brr
    Sheet22.Activate
    
    Application.ScreenUpdating = True
    'MsgBox Format(Timer - t, "0.00" & "s")
    
End Sub