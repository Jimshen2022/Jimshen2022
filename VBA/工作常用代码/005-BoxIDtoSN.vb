Sub BoxIDtoSN()

Application.ScreenUpdating = False
'On Error Resume Next
Dim ar_Box As Variant
Dim i As Long
Dim ar_serial As Variant
With Sheet7
    ar_Box = .Range("A1").CurrentRegion
    
    ReDim ar_serial(1 To 10 * UBound(ar_Box), 1 To 2)
    k = 1
    For i = 2 To UBound(ar_Box)
 
       '当流水号位数小于5位数时，直接以当前value去填数组
        If Len(ar_Box(i, 2)) < 5 Then
            k = k + 1
            ar_serial(k, 1) = "'" & ar_Box(i, 1)
            ar_serial(k, 2) = ar_Box(i, 2) * 1
        '如果SN以W开头，则执行以下split 将SN拆分（不含W or w)，并以“/”后数字循环产生SN
        ElseIf Mid(ar_Box(i, 2), 1, 1) = "W" Or Mid(ar_Box(i, 2), 1, 1) = "w" Then
            For j = 0 To Val(Split(Mid(ar_Box(i, 2), 2, 13), "/")(1)) - 1
            k = k + 1
            ar_serial(k, 1) = "'" & ar_Box(i, 1)
            ar_serial(k, 2) = "'" & Split(Mid(ar_Box(i, 2), 2, 13), "/")(0) + j
            Next
        '除以上条件，则执行以下Split, 将SN拆分，并以“/”后数字循环产生SN
        Else
        
            For j = 0 To Val(Split(ar_Box(i, 2), "/")(1)) - 1
            k = k + 1
            ar_serial(k, 1) = "'" & ar_Box(i, 1)
            ar_serial(k, 2) = "'" & Split(ar_Box(i, 2), "/")(0) + j
            
            Next
        End If
    Next
    ar_serial(1, 1) = "ITEM"
    ar_serial(1, 2) = "SERIAL NUMBER"
    Sheet9.Cells.Clear
    Sheet9.Columns("a:b").NumberFormat = "@"
    Sheet9.Range("a1").Resize(UBound(ar_serial), UBound(ar_serial, 2)) = ar_serial
    Sheet9.Columns("a:b").AutoFit
End With

Application.ScreenUpdating = True
End Sub