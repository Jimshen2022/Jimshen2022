Sub STO_columnA_product() 'Column A product ----finished
    
    Dim i&, nrow&, arr()
    nrow = Sheet2.Range("p1048576").End(3).Row
    arr = Sheet2.Range("a1").CurrentRegion
    For i = 2 To nrow
        If Left(arr(i, 16), 4) = "100-" Or Left(arr(i, 16), 1) = "A" Or Left(arr(i, 16), 1) = "B" Or Left(arr(i, 16), 1) = "D" Or Left(arr(i, 16), 1) = "H" Or Left(arr(i, 16), 1) = "L" Or Left(arr(i, 16), 1) = "P" Or Left(arr(i, 16), 1) = "Q" Or Left(arr(i, 16), 1) = "T" Or Left(arr(i, 16), 1) = "R" Or Left(arr(i, 16), 1) = "W" Or Left(arr(i, 16), 1) = "Z" Then
            arr(i, 1) = "CG"
        Else: arr(i, 1) = "UPH"
        End If
    Next
    Sheet2.Range("a1").Resize(UBound(arr), 33) = arr
    
End Sub