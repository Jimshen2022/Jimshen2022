Sub PullWanek3STO()  '¿ç¹¤×÷±¡ÌáÈ¡ÄÚÈÝ
    
    Application.ScreenUpdating = False
    Dim wb As Workbook
    Dim arr
    
    Sheet3.Activate
    Columns("e:v").ClearContents
    
    Set wb = GetObject("C:\Users\jishen\Downloads\WANEK3STO.xlsx")      '´ò¿ª¹¤×÷²¾
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False

    Sheet3.Columns("e:v").NumberFormat = "@"
    Sheet3.Range("e1").Resize(UBound(arr), 18) = arr

    Erase arr
    Application.ScreenUpdating = True
    
End Sub

Sub Pull_Wanek3_DC_361_OUT_trx()  '¿ç¹¤×÷±¡ÌáÈ¡ÄÚÈÝ
    
    Application.ScreenUpdating = False
    
    Dim wb As Workbook
    Dim arr
    Sheet5.Activate
    Columns("i1:aj1048576").ClearContents
    
    Set wb = GetObject("C:\Users\jishen\Downloads\WANEK3DC361OUT.xlsx")      '´ò¿ª¹¤×÷²¾
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False

    Sheet5.Columns("i:aj").NumberFormat = "@"
    Sheet5.Range("i1").Resize(UBound(arr), 28) = arr
    Erase arr
    Application.ScreenUpdating = True
    
End Sub

Sub Pull_Wanek3_111_in_trx()  '¿ç¹¤×÷±¡ÌáÈ¡ÄÚÈÝ
    
    Application.ScreenUpdating = False
    Dim wb As Workbook
    Dim arr
    Sheet1.Activate
    Columns("e:af").ClearContents
    Set wb = GetObject("C:\Users\jishen\Downloads\WANEK3111IN.xlsx")      '´ò¿ª¹¤×÷²¾
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False
    Sheet1.Columns("e:af").NumberFormat = "@"
    Sheet1.Range("e1").Resize(UBound(arr), 28) = arr

    Erase arr
    Application.ScreenUpdating = True

End Sub


Sub Pull_Wanek3_202_out_trx()  '¿ç¹¤×÷±¡ÌáÈ¡ÄÚÈÝ
    
    Application.ScreenUpdating = False
    
    Dim wb As Workbook
    Dim arr, brr, i&, j&
    Sheet2.Activate
    Columns("e:az").ClearContents
    
    Set wb = GetObject("C:\Users\jishen\Downloads\WANEK3202OUT.xlsx")      '´ò¿ª¹¤×÷²¾
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False
    ReDim brr(1 To UBound(arr), 1 To 28)
    For i = 1 To UBound(arr)
            If Not arr(i, 12) Like "C*" And Not arr(i, 12) Like "U*" And Not arr(i, 12) Like "M*" And Not arr(i, 12) Like "B1*" And Not arr(i, 12) Like "B2*" And Not arr(i, 12) Like "B6*" Then
                m = m + 1
                For j = 1 To 28
                    brr(m, j) = arr(i, j)
                Next
            End If
    Next
    Sheet2.Columns("e:af").NumberFormat = "@"
    Sheet2.Range("e1").Resize(UBound(arr), 28) = brr
    Erase arr
    Erase brr
    Application.ScreenUpdating = True

End Sub


Sub Pull_DC_202_in_trx()  '¿ç¹¤×÷±¡ÌáÈ¡ÄÚÈÝ

    Application.ScreenUpdating = False
    Dim wb As Workbook
    Dim arr, brr, i&, j&

    Sheet4.Activate
    Columns("e:af").ClearContents
    
    Set wb = GetObject("C:\Users\jishen\Downloads\DC202IN.xlsx")      '´ò¿ª¹¤×÷²¾
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False
    ReDim brr(1 To UBound(arr), 1 To 28)
    For i = 1 To UBound(arr)
            If arr(i, 9) Like "UL6*" Or arr(i, 9) Like "M*" Or arr(i, 9) = "To Location ID" Then
                m = m + 1
                For j = 1 To 28
                    brr(m, j) = arr(i, j)
                Next
            End If
    Next
    Sheet4.Columns("e:af").NumberFormat = "@"
    Sheet4.Range("e1").Resize(UBound(arr), 28) = brr
       
    Erase arr
    Erase brr
    Application.ScreenUpdating = True
    
End Sub

Sub Pull_BW_202_in_trx()  '¿ç¹¤×÷±¡ÌáÈ¡ÄÚÈÝ

    Application.ScreenUpdating = False
    Dim wb As Workbook
    Dim arr, brr, i&, j&

    Sheet12.Activate
    Columns("e:af").ClearContents
    
    Set wb = GetObject("C:\Users\jishen\Downloads\DC202IN.xlsx")      '´ò¿ª¹¤×÷²¾
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False
    ReDim brr(1 To UBound(arr), 1 To 28)
    For i = 1 To UBound(arr)
            If arr(i, 9) Like "UL9*" Or arr(i, 9) Like "B1*" Or arr(i, 9) Like "B2*" Or arr(i, 9) = "To Location ID" Then
                m = m + 1
                For j = 1 To 28
                    brr(m, j) = arr(i, j)
                Next
            End If
    Next
    Sheet12.Columns("e:af").NumberFormat = "@"
    Sheet12.Range("e1").Resize(UBound(arr), 28) = brr
       
    Erase arr
    Erase brr
    Application.ScreenUpdating = True
    
End Sub


Sub Pull_BW_202_out_trx()  '¿ç¹¤×÷±¡ÌáÈ¡ÄÚÈÝ
    
    Application.ScreenUpdating = False
    
    Dim wb As Workbook
    Dim arr, brr, i&, j&
    Sheet6.Activate

    Range("a2:az1048576").ClearContents
    
    Set wb = GetObject("C:\Users\jishen\Downloads\WANEK3202OUT.xlsx")      '´ò¿ª¹¤×÷²¾
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False
    ReDim brr(1 To UBound(arr), 1 To 28)
    For i = 1 To UBound(arr)
            If (arr(i, 12) Like "B6*" Or arr(i, 12) Like "Ref*") And (arr(i, 9) Like "CM*" Or arr(i, 9) Like "To Location*") Then
                m = m + 1
                For j = 1 To 28
                    brr(m, j) = arr(i, j)
                Next
            End If
    Next
    Sheet6.Columns("e:af").NumberFormat = "@"
    Sheet6.Range("e1").Resize(UBound(arr), 28) = brr
    Erase arr
    Erase brr

' QUERY UP FOR COLUMNS B
Dim d As Object, arrr, brrr, x&
row = Sheet6.Range("e1048576").End(3).row


Set d = CreateObject("scripting.dictionary")
    d.CompareMode = vbTextCompare '²»Çø·Ö×ÖÄ¸´óÐ¡Ð´
    
arrr = Sheet7.Range("a1").CurrentRegion  'Êý¾ÝÔ´×°ÈëÊý×éarr
brrr = Sheet6.Range("a1").CurrentRegion  '²éÑ¯ÇøÓòÊý¾Ý×°ÈëÊý×ébrr

For x = 1 To UBound(arrr)            '±éÀúÊý×éarr
    d(arrr(x, 1)) = arrr(x, 2)        '½«item + up ×÷Îªkey£¬×°Èë×Öµä
Next

For x = 2 To UBound(brrr)            '±êÌâÐÐ²»ÓÃ²éÑ¯£¬ËùÒÔ´ÓµÚ¶þÐÐ¿ªÊ¼±éÀú²éÑ¯ÊýÖµbrr
    If d.exists(brrr(x, 5)) Then brrr(x, 2) = d(brrr(x, 5)) Else brrr(x, 2) = 200
    'Èç¹û×ÖµäÖÐ´æÔÚitem, '¸ù¾Ýitem ´Ó×ÖµäÖÐÈ¡UPÖµ
    brrr(x, 1) = brrr(x, 2) * brrr(x, 3)
Next x

 'Sheet5.[a1].Resize(UBound(brr), UBound(brr, 2)) = brr   '(brr,2)ÊÇÖ¸Êý×éµÄ¶þÎ³ÏÂ±ê
With Sheet6.Range("a1").CurrentRegion

    .Value = brrr
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter

End With
    Set d = Nothing        'ÊÍ·Å×Öµä
    Erase arrr
    Erase brrr
    

'Qty and date update in columns C and D

Dim crr
crr = Sheet6.Range("a1").CurrentRegion
row = Sheet6.Range("e1048576").End(3).row


Dim z&
For z = 2 To UBound(crr)
    
    crr(z, 3) = crr(z, 18) * 1
    crr(z, 1) = crr(z, 3) * crr(z, 2)
    crr(z, 4) = CDate(crr(z, 24))

    Next z
    
    With Sheet6.Range("a1").CurrentRegion

    .Value = crr
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    Range("d2:d" & row).NumberFormat = "mm-dd-yyyy"
    
    End With
    Erase crr
    


    
    Application.ScreenUpdating = True

End Sub


