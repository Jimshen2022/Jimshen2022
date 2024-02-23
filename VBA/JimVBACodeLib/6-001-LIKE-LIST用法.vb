Sub MIL_368_TRX()  'FG_LoadingScanned sheet

    Application.ScreenUpdating = False
    Call Filter666
    't = Timer
    Dim wb As Workbook
    Dim arr, brr, i&, j&, nrow&
    
    Sheet8.Activate
    With Sheet8
        .Cells.ClearContents
        
        Set wb = GetObject("C:\Users\jishen\Downloads\MIL-368.xlsx")      '打开工作簿
        nrow = wb.ActiveSheet.Range("a1048576").End(3).Row
        arr = wb.ActiveSheet.Range("a1:ac" & nrow)
        wb.Close False
        
        .Columns("d:af").NumberFormat = "@"
        .Range("d1").Resize(UBound(arr), 29) = arr
        .Range("a1:c1").Value = Array("Product", "Hour", "Date")

    
        brr = .Range("a1").CurrentRegion
        
        For i = 2 To UBound(brr)
            'Product
            If brr(i, 4) Like "E*" Then
                brr(i, 1) = "CG"
            ElseIf brr(i, 4) Like "M*" And brr(i, 5) = "ZKIS" Then
                brr(i, 1) = "Bedding"
            ElseIf brr(i, 4) Like "[0-9 U]*" Then
                brr(i, 1) = "UPH"
            Else
                brr(i, 1) = "Check"
            End If
            'Hour
            brr(i, 2) = Mid(brr(i, 24), 1, 2)
            'Date
            brr(i, 3) = CDate(brr(i, 23))
            
        Next
        
        .Columns("d:af").NumberFormat = "@"
        .Range("a1").Resize(UBound(brr), 32) = brr
        .Range("a:af").Columns.AutoFit
    End With
    
    Erase arr
    Erase brr
    Application.ScreenUpdating = True
    'MsgBox "Data Downloaded Successful!  " & Format(Timer - t, "0.00" & "s")
End Sub