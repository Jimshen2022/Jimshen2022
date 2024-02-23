

'Judge product by item number first chart.
Sub STO_columnA_product()

Dim i&, nrow&, arr()
nrow = Sheet4.Range("q1048576").End(3).Row
arr = Sheet4.Range("a1").CurrentRegion
    For i = 2 To nrow
        If Left(arr(i, 17), 4) = "100-" Or (Left(arr(i, 17), 1) Like "[ABDEHLMQRTWXZ]") Then
            arr(i, 1) = "CG"
        ElseIf Left(arr(i, 10), 1) <> "Z" Then arr(i, 1) = "RP"
        Else: arr(i, 1) = "UPH"
        End If
    Next
Sheet4.Range("a1").Resize(UBound(arr), 33) = arr
 
End Sub





Sub product_separate() 'Asia_OnHand - 20210221.1309.xlsx
    Application.ScreenUpdating = False
    Dim i&, nrow&, Arr()
    With Sheet11
        nrow =  .Range("a1048576").End(3).Row
         .Range("j1") = "Product"
        Arr =  .Range("a1").CurrentRegion
        For i = 2 To UBound(Arr)
            Select Case Mid(Arr(i, 1), 1, 1)
                Case Is = "A"
                    Arr(i, 10) = "Accessary"
                Case Is = "B"
                    Arr(i, 10) = "CG"
                Case Is = "D"
                    Arr(i, 10) = "CG"
                Case Is = "H"
                    Arr(i, 10) = "CG"
                Case Is = "L"
                    Arr(i, 10) = "Accessary"
                Case Is = "M"
                    Arr(i, 10) = "UPH"
                Case Is = "P"
                    Arr(i, 10) = "CG"
                Case Is = "Q"
                    Arr(i, 10) = "Accessary"
                Case Is = "R"
                    Arr(i, 10) = "Accessary"
                Case Is = "T"
                    Arr(i, 10) = "CG"
                Case Is = "W"
                    Arr(i, 10) = "CG"
                Case Is = "Z"
                    Arr(i, 10) = "CG"
                Case Else
                    Arr(i, 10) = "UPH"
                    
            End Select
        Next
        
         .Columns("a:a").NumberFormat = "@"
         .Range("a1").Resize(UBound(Arr), UBound(Arr, 2)).Value = Arr
        
    End With
    
    Application.ScreenUpdating = True
End Sub


