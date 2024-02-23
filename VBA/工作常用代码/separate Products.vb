Sub sheet1ColumnN()
    
    Application.ScreenUpdating = False
    Dim i&, nrow&, arr()
    Sheet1.Range("n1:t1") = Array("Date", "Month", "Container#", "ContainerQty", "Product", "M&S", "Type")
    
    With Sheet1
        nrow =  .Range("a1048576").End(3).Row
        arr =  .Range("a1").CurrentRegion
        For i = 2 To nrow
            arr(i, 14) = CDate(Mid(arr(i, 7), 1, 10))
            arr(i, 15) = Month(arr(i, 14))
            arr(i, 16) = arr(i, 1) & arr(i, 2) & arr(i, 7)
            If arr(i, 1) = arr(i - 1, 1) Then
                arr(i, 17) = 0
            Else
                arr(i, 17) = 1
            End If
            If arr(i, 10) = "WPLS" And Not arr(i, 5) Like "[a-z|A-Z]" Then
                arr(i, 18) = "Plastic"
            ElseIf Mid(arr(i, 10), 1, 1) <> "Z" Then
                arr(i, 18) = "Raw Materials"
            ElseIf Mid(arr(i, 10), 1, 1) = "Z" And Right(arr(i, 10), 1) = "K" Then
                arr(i, 18) = "UnKits"
            ElseIf Mid(arr(i, 10), 1, 1) = "Z" And Right(arr(i, 10), 1) = "Z" Then
                arr(i, 18) = "ZipperCover"
            ElseIf Mid(arr(i, 10), 1, 1) = "Z" And Right(arr(i, 10), 1) = "S" Then
                arr(i, 18) = "Bedding"
            ElseIf Mid(arr(i, 10), 1, 1) = "Z" And (Right(arr(i, 10), 1) = "B" Or Right(arr(i, 10), 1) = "W") Then
                arr(i, 18) = "CG"
            Else
                arr(i, 18) = "UPH"
            End If
            
            
            '=IF(AND(EXACT(LOWER(B3),UPPER(B3)),C3="WPLS"),"Plastic", & _
            If (Mid(C3, 1, 1) <> "Z", "Raw Materials",  &  _
                        If ( And (Mid(C3, 1, 1) = "Z", Right(C3, 1) = "K"), "Un-Kits",  &  _
                            If ( And (Mid(C3, 1, 1) = "Z", Right(C3, 1) = "Z"), "Zipper Cover",  &  _
                                If ( And (Mid(C3, 1, 1) = "Z", Right(C3, 1) = "S"), "Bedding",  &  _
                                    If ( And (Mid(C3, 1, 1) = "Z",  Or (Right(C3, 1) = "W", Right(C3, 1) = "B")), "CG", "UPH"))))))
                                
                            Next
                             .Columns("a:e").NumberFormat = "@"
                             .Columns("g:h").NumberFormat = "@"
                             .Range("a1").Resize(UBound(arr), UBound(arr, 2)).Value = arr
                             .Columns("A:T").AutoFit
                        End With
                        Erase arr
                        
                        Application.ScreenUpdating = True
                        
                    End Sub
                    
                    
                    
                    
