' DICTIONARY SUMMARIZED ITEM AND QTY
' BW,MPT INVENTORY TURNOVER RATIO FILES 

Sub Inbound_Amt_Qty()
    
    Application.ScreenUpdating = False
    't = Timer
    Dim arr, brr, crr, i&, j&, k&, arow&, brow&, crow&
    Dim ad1 As Object, ad2 As Object, bd1 As Object, bd2 As Object, cd1 As Object, cd2 As Object
    

    arow = Sheet1.Range("e1048576").End(3).row
    brow = Sheet4.Range("e1048576").End(3).row
    crow = Sheet12.Range("e1048576").End(3).row
    
    'WANEK3_IN
    arr = Sheet1.Range("a1:e" & arow)
    For i = 2 To UBound(arr)
        arr(1, 5) = "Whse"
        arr(i, 5) = "Wanek3"
    Next

    'DC_IN
    brr = Sheet4.Range("a1:e" & brow)
    For j = 2 To UBound(brr)
        brr(1, 5) = "Whse"
        brr(j, 5) = "DC"
        
    Next
    
    'BW_202_OUT
    crr = Sheet12.Range("a1:e" & crow)
    For k = 2 To UBound(crr)
        crr(1, 5) = "Whse"
        crr(k, 5) = "BW"
    Next
    
    
    Set ad1 = CreateObject("scripting.dictionary")
    Set ad2 = CreateObject("scripting.dictionary")
    Set bd1 = CreateObject("scripting.dictionary")
    Set bd2 = CreateObject("scripting.dictionary")
    Set cd1 = CreateObject("scripting.dictionary")
    Set cd2 = CreateObject("scripting.dictionary")
    ad1.CompareMode = vbTextCompare
    ad2.CompareMode = vbTextCompare
    bd1.CompareMode = vbTextCompare
    bd2.CompareMode = vbTextCompare
    cd1.CompareMode = vbTextCompare
    cd2.CompareMode = vbTextCompare
    
 'WANEK3_IN loading into dic
    For i = 2 To UBound(arr)
        If arr(i, 4) >= (Date - 30) Then
            ad1(arr(i, 5)) = ad1(arr(i, 5)) + arr(i, 1)  'Site + Amt
            ad2(arr(i, 5)) = ad2(arr(i, 5)) + arr(i, 3)  'Site + Qty
        End If
    Next
    
  'DC_IN loading into dic  db1 and db2
    For i = 2 To UBound(brr)
        If brr(i, 4) >= Date - 30 Then
            bd1(brr(i, 5)) = bd1(brr(i, 5)) + brr(i, 1)  'Site + Amt
            bd2(brr(i, 5)) = bd2(brr(i, 5)) + brr(i, 3)  'Site + Qty
        End If
    Next
    
  'BW_IN loading into dic  cb1 and cb2
    For i = 2 To UBound(crr)
        If crr(i, 4) >= Date - 30 Then
            cd1(crr(i, 5)) = cd1(crr(i, 5)) + crr(i, 1)  'Site + Amt
            cd2(crr(i, 5)) = cd2(crr(i, 5)) + crr(i, 3)  'Site + Qty
        End If
    Next
                  
    
'Sheet InvTurnOverRatio fill out Ending_Inv_Amt and Qty
    With Sheet13

'amt
        .Cells(2, "C").Value = cd1("BW")
        .Cells(3, "C").Value = bd1("DC")
        .Cells(4, "C").Value = cd1("BW") + bd1("DC")
'qty
        .Cells(10, "C").Value = cd2("BW")
        .Cells(11, "C").Value = bd2("DC")
        .Cells(12, "C").Value = cd2("BW") + bd2("DC")

'Beginning_Inv_Amt_Qty

        .Range("b2").Value = .Range("e2") + .Range("d2") - .Range("c2")
        .Range("b3").Value = .Range("e3") + .Range("d3") - .Range("c3")
        .Range("b4").Value = .Range("e4") + .Range("d4") - .Range("c4")
        .Range("b10").Value = .Range("e10") + .Range("d10") - .Range("c10")
        .Range("b11").Value = .Range("e11") + .Range("d11") - .Range("c11")
        .Range("b12").Value = .Range("e12") + .Range("d12") - .Range("c12")
        .Range("f2") = .Range("d2") / ((.Range("b2") + .Range("e2")) / 2) * 12
        .Range("f3") = .Range("d3") / ((.Range("b3") + .Range("e3")) / 2) * 12
        .Range("f4") = .Range("d4") / ((.Range("b4") + .Range("e4")) / 2) * 12
        .Range("b2:e12").NumberFormat = "##,##0"
        .Range("f2:f4").NumberFormat = "##,##0.0"
    End With
    
    Set ad1 = Nothing
    Set ad2 = Nothing
    Set bd1 = Nothing
    Set bd2 = Nothing
    Set cd1 = Nothing
    Set cd2 = Nothing
    Erase arr
    Erase brr
    Erase crr
    'ThisWorkbook.Save
    Application.ScreenUpdating = True
    'MsgBox Format(Timer - t, "0.00" & "s")

    
End Sub