
'MIL Inventory Turns - 2022.xlsx
Sub AmtInOut()
    'Turns by Amt sheet B~W IN AND OUT Calculation
    
    Application.ScreenUpdating = False
    Dim arr, brr, i&, j&, k&, m&, x&, l&, d As Object, d2 As Object
    
    arr = Sheet3.Range("a1").CurrentRegion
    Set d = CreateObject("scripting.dictionary")
    Set d2 = CreateObject("scripting.dictionary")
    d.CompareMode = vbTextCompare
    d2.CompareMode = vbTextCompare
   'load Trx into Dic,  YearWeek/Product/INOUT as Key,  AMT as items
    For m = 2 To UBound(arr)
        d(arr(m, 6) & "/" & arr(m, 9) & "/" & arr(m, 8)) = d(arr(m, 6) & "/" & arr(m, 9) & "/" & arr(m, 8)) + arr(m, 11)
    Next
    
    For l = 2 To UBound(arr)
        d2(arr(l, 6)) = ""
    Next
    
    Sheet8.Range("a3:bd1000").ClearContents
    Sheet8.Range("a3").Resize(d2.Count, 1).Value = Application.Transpose(d2.keys)
    
    
    Sheet8.Activate
    brr = Sheet8.Range("a1").CurrentRegion
    'calculate B-BW,  by YearWeek/Product
    For i = 3 To UBound(brr)
        For j = 2 To UBound(brr, 2)
             brr(i, j) = d(brr(i, 1) & "/" & brr(1, j) & "/" & brr(2, j))
        Next
    Next
    Sheet8.Range("a1").Resize(UBound(brr), UBound(brr, 2)).Value = brr
    Erase arr, brr
    Set d = Nothing
    Set d2 = Nothing
    
    Application.ScreenUpdating = True
End Sub