Sub UnitCubesRanges()
    Application.ScreenUpdating = False
    Dim i, j, arr, brr


    Sheet2.Range("ag1:ah1048576").ClearContents
    Sheet2.Range("ag1").Value = "UnitCubeRange"
    Sheet2.Range("ah1").Value = "Product2"
        
    brr = Sheet10.Range("a1").CurrentRegion
    arr = Sheet2.Range("a1").CurrentRegion
    
    With Sheet2
        
        For j = 2 To UBound(arr)
            If arr(j, 3) < 0.3 Then
                arr(j, 33) = "0-0.3"
            ElseIf arr(j, 3) < 0.5 Then
                arr(j, 33) = "0.3-0.5"
            ElseIf arr(j, 3) < 0.7 Then
                arr(j, 33) = "0.5-0.7"
            ElseIf arr(j, 3) < 0.9 Then
                arr(j, 33) = "0.7-0.9"
            ElseIf arr(j, 3) < 1.1 Then
                arr(j, 33) = "0.9-1.1"
            ElseIf arr(j, 3) < 1.3 Then
                arr(j, 33) = "1.1-1.3"
            ElseIf arr(j, 3) < 1.5 Then
                arr(j, 33) = "1.3-1.5"
            ElseIf arr(j, 3) < 1.7 Then
                arr(j, 33) = "1.5-1.7"
            ElseIf arr(j, 3) < 1.9 Then
                arr(j, 33) = "1.7-1.9"
            ElseIf arr(j, 3) < 2.1 Then
                arr(j, 33) = "1.9-2.1"
            ElseIf arr(j, 3) < 2.3 Then
                arr(j, 33) = "2.1-2.3"
            ElseIf arr(j, 3) < 2.5 Then
                arr(j, 33) = "2.3-2.5"
            ElseIf arr(j, 3) < 2.7 Then
                arr(j, 33) = "2.5-2.7"
            ElseIf arr(j, 3) < 2.9 Then
                arr(j, 33) = "2.7-2.9"
            Else
                arr(j, 33) = "Over 2.9"
            End If
        
            If Mid(arr(j, 16), 1, 1) Like "[ALQR]" Then
                arr(j, 34) = "Accessory"
            ElseIf Mid(arr(j, 16), 1, 1) Like "[BDEHPTWZ]" Then
                arr(j, 34) = "CG"
            ElseIf Mid(arr(j, 16), 1, 1) Like "M*" Then
                arr(j, 34) = "Mattress"
            ElseIf Mid(arr(j, 16), 1, 4) Like "100-*" Then
                arr(j, 34) = "CG"
            Else
                arr(j, 34) = "UPH"
            End If
            
        Next
        .Range("ag1:ag" & UBound(arr)) = Application.Index(arr, , 33)
        .Range("ah1:ah" & UBound(arr)) = Application.Index(arr, , 34)
    Erase arr, brr
    
    End With
    Application.ScreenUpdating = True


End Sub