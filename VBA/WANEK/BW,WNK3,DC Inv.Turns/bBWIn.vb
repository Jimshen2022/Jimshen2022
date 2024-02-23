
Sub BW_IN()  'for BW_IN column C and D

't = Timer
Application.ScreenUpdating = False
Sheet12.Activate
Sheet12.Range("a2:d1048576").ClearContents

Dim arr, brr(1 To 1048576, 1 To 2), row&
arr = Sheet12.Range("a1").CurrentRegion
row = Sheet12.Range("e1048576").End(3).row


Dim x&
For x = 2 To UBound(arr)
    arr(x, 3) = arr(x, 18) * 1
    arr(x, 4) = CDate(arr(x, 24))

    Next x
    
    With Sheet12.Range("a1").CurrentRegion

    .Value = arr
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    Range("d2:d" & row).NumberFormat = "mm-dd-yyyy"
    End With
    Erase arr
   
Application.ScreenUpdating = True
    'MsgBox "udpated~ " & Timer - t
End Sub

Sub UPfind_in_BW_IN()  'ÒýÓÃUnite Price  'for BW_IN column B

't = Timer
Application.ScreenUpdating = False
Sheet12.Activate
Dim d As Object, arr, brr, i&
row = Sheet12.Range("e1048576").End(3).row


Set d = CreateObject("scripting.dictionary")
    d.CompareMode = vbTextCompare '²»Çø·Ö×ÖÄ¸´óÐ¡Ð´
    
arr = Sheet7.Range("a1").CurrentRegion  'Êý¾ÝÔ´×°ÈëÊý×éarr
brr = Sheet12.Range("a1").CurrentRegion  '²éÑ¯ÇøÓòÊý¾Ý×°ÈëÊý×ébrr

For i = 1 To UBound(arr)            '±éÀúÊý×éarr
    d(arr(i, 1)) = arr(i, 2)        '½«item + up ×÷Îªkey£¬×°Èë×Öµä
Next

For i = 2 To UBound(brr)            '±êÌâÐÐ²»ÓÃ²éÑ¯£¬ËùÒÔ´ÓµÚ¶þÐÐ¿ªÊ¼±éÀú²éÑ¯ÊýÖµbrr
    If d.exists(brr(i, 5)) Then brr(i, 2) = d(brr(i, 5)) Else brr(i, 2) = 200
    'Èç¹û×ÖµäÖÐ´æÔÚitem, '¸ù¾Ýitem ´Ó×ÖµäÖÐÈ¡UPÖµ
    brr(i, 1) = brr(i, 2) * brr(i, 3)
Next i

 'Sheet5.[a1].Resize(UBound(brr), UBound(brr, 2)) = brr   '(brr,2)ÊÇÖ¸Êý×éµÄ¶þÎ³ÏÂ±ê
With Sheet12.Range("a1").CurrentRegion

    .Value = brr
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter

End With
    Set d = Nothing        'ÊÍ·Å×Öµä
    Erase arr

Application.ScreenUpdating = True
   ' MsgBox "udpated~ " & Timer - t
End Sub
