Sub AddColumnAFor10200126()

    Application.ScreenUpdating = False
    Dim i&, arr()
    Sheet2.Range("a1") = "Location"

    
    arr = Sheet2.Range("a1").CurrentRegion
        For i = 2 To UBound(arr)
            If Len(arr(i, 4)) = 1 Then
                arr(i, 1) = arr(i, 3) & "000" & arr(i, 4) & arr(i, 5) & arr(i, 6)
            ElseIf Len(arr(i, 4)) = 2 Then
                    arr(i, 1) = arr(i, 3) & "00" & arr(i, 4) & arr(i, 5) & arr(i, 6)
            Else
                    arr(i, 1) = arr(i, 3) & arr(i, 4) & arr(i, 5) & arr(i, 6)
            End If
            
        Next
    Sheet2.Columns("a:g").NumberFormat = "@"
    Sheet2.Range("a1").Resize(UBound(arr), UBound(arr, 2)).Value = arr
    Sheet2.Columns("a:k").AutoFit
    
    Erase arr
    
    Application.ScreenUpdating = True
End Sub


Sub Sheet3ColumnHtoL()

    Application.ScreenUpdating = False
    Dim d As Object, arr, brr, crr, i&
    Set d = CreateObject("scripting.dictionary")
    Set d2 = CreateObject("scripting.dictionary")
    
    d.CompareMode = vbTextCompare                    '²»Çø·Ö×ÖÄ¸´óÐ¡Ð´
    arr = Sheets("Yard").Range("a1").CurrentRegion   'Êý¾ÝÔ´×°ÈëÊý×éarr
    Sheet3.Range("h1:L1") = Array("Yard", "1020.01.26", "Balance", "Type", "ABS(Variance)")
    brr = Sheets("AS400vs1020").Range("a1").CurrentRegion       '²éÑ¯ÇøÓòÊý¾Ý×°ÈëÊý×ébrr
    crr = Sheets("1020.01.26").Range("a1").CurrentRegion
    
    For i = 1 To UBound(arr)
        d(arr(i, 11)) = d(arr(i, 11)) + arr(i, 14)       '½«ITEM¼°QTYÀÛ¼Ó ×°Èë×Öµä
        Next
    
    For i = 1 To UBound(crr)
        d2(crr(i, 7)) = d2(crr(i, 7)) + crr(i, 9)
        Next
        
    For i = 2 To UBound(brr)
        If d.exists(brr(i, 1)) Then
            brr(i, 8) = d(brr(i, 1))   'Èç¹û×ÖµäÖÐÓÐÖµ£¬Ôò·µ»ØQty¼Ó×Ü
        Else
            brr(i, 8) = 0  'Èç¹û×ÖµäÖÐ²»´æÔÚ£¬ÔòÖµ·µ»ØÎª0
        End If
    Next
    
    
    For i = 2 To UBound(brr)
        If d2.exists(brr(i, 1)) Then
            brr(i, 9) = d2(brr(i, 1))   'Èç¹û×ÖµäÖÐÓÐÖµ£¬Ôò·µ»ØQty¼Ó×Ü
        Else
            brr(i, 9) = 0  'Èç¹û×ÖµäÖÐ²»´æÔÚ£¬ÔòÖµ·µ»ØÎª0
        End If
        brr(i, 10) = brr(i, 8) + brr(i, 9) - brr(i, 4)
        brr(i, 12) = Abs(brr(i, 10))
        
        If brr(i, 10) > 0 Then
            brr(i, 11) = "¿â´æÅÌÓ¯"
        ElseIf brr(i, 10) = 0 Then
            brr(i, 11) = "Ã»ÓÐ²îÒì"
        Else
            brr(i, 11) = "¿â´æÅÌ¿÷"
        End If
    Next
    
    
    
    
    With Sheet3
        .Columns("a:c").NumberFormat = "@"
        .Range("a1").Resize(UBound(brr), UBound(brr, 2)).Value = brr
        

    End With

    Set d = Nothing
    Erase arr
    Erase brr
    Application.ScreenUpdating = True
    
End Sub

