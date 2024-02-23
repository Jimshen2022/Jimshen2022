
Sub load_UP_from_as400()  'sheets UP
    
    Application.ScreenUpdating = False
    Dim i As Long
    Dim adors As New Recordset
    
    
    Sheet7.Activate
    Sheet7.Cells.Clear
    'Sheet14.Range("a1").CurrentRegion.Copy Sheet7.Range("a1")
    
    Set Db = New Connection
    Db.CursorLocation = adUseClient
    If Db.State = 1 Then Db.Close
    U = Sheet9.Range("a1").Value
    P = Sheet9.Range("a2").Value
    
   
    Db.Open "Provider =IBMDASQL.DataSource.1" & _
     ";Catalog Library List=JIMTDTA" & _
     ";Persist Security Info=True" & _
     ";Force Translate=0" & _
     ";Data Source = 10.9.3.106 " & _
     ";User ID =" & U & "" & _
     ";Password =" & P
          
     Set adors = New Recordset
     If adors.State = 1 Then adors.Close

'    cmdtxt = "SELECT u.ITNBR,u.AVCST/23081 as UP_$USD  " & _
'             "FROM AMFLIBW.ITEMBL u " & _
'             "Where u.HOUSE IN ('35') And u.AVCST<>0 AND u.ITCLS like 'Z%' AND u.ITCLS not like '%K' "
             
        cmdtxt = "SELECT a.RPAITX,(CASE WHEN a.RPBRCD IN ('VND') THEN a.RPAMVA/22685 ELSE a.RPAMVA END) AS RPAMVA,a.RPBLDT,a.RPZ0D7, T2.ITCLS " & _
             "FROM AMFLIBW.ITMFPR a " & _
             "left join AMFLIBW.ITMRVA T2 on a.RPAITX=T2.ITNBR and a.RPZ0D7 = T2.STID " & _
             "WHERE a.RPZ0D7 = '35' AND a.RPAITX||a.RPZ0D7||a.RPBLDT IN " & _
             "(SELECT a.RPAITX||a.RPZ0D7||MAX(a.RPBLDT) RPBLDT " & _
             "FROM AMFLIBW.ITMFPR a  WHERE a.RPZ0D7 = '35' GROUP BY a.RPAITX,a.RPZ0D7) " & _
             "AND T2.ITCLS LIKE 'Z%' "
             

    adors.Open cmdtxt, Db, 3, 3
  
     For i = 0 To adors.Fields.Count - 1
         Sheet7.Cells(1, i + 1) = adors.Fields(i).Name
     Next i
     
     Sheet7.Range("a1048576").End(3).Offset(1, 0).CopyFromRecordset adors
     adors.Close
     Set adors = Nothing
    Sheet7.Columns("A:A").NumberFormat = "@"
    Application.ScreenUpdating = True
    
End Sub
Sub load_UP2_from_as400()  'sheets UP
    
    Application.ScreenUpdating = False
    Dim i As Long
    Dim adors As New Recordset
    Sheet14.Activate
    Cells.Clear
    Sheet14.Range("a2:e1048576").ClearContents
    Sheet14.Columns("A:A").NumberFormat = "@"
    
    Set Db = New Connection
    Db.CursorLocation = adUseClient
    If Db.State = 1 Then Db.Close
    U = Sheet9.Range("a1").Value
    P = Sheet9.Range("a2").Value
   
    Db.Open "Provider =IBMDASQL.DataSource.1" & _
     ";Catalog Library List=JIMTDTA" & _
     ";Persist Security Info=True" & _
     ";Force Translate=0" & _
     ";Data Source = 10.9.3.106" & _
     ";User ID =" & U & "" & _
     ";Password =" & P
     
     Set adors = New Recordset
     If adors.State = 1 Then adors.Close

    cmdtxt = "SELECT T1.RPAITX,T1.RPAMVA,T1.RPZ0D7,T2.ITCLS,T1.RPBLDT " & _
             "FROM AMFLIBW.ITMFPR T1 left join AMFLIBW.ITMRVA T2 on T1.RPAITX=T2.ITNBR and T1.RPZ0D7 = T2.STID " & _
             "WHERE T1.RPZ0D7 = '35'" & _
             "Order by T1.RPAITX,T1.RPBLDT desc "
             
    adors.Open cmdtxt, Db, 3, 3
  
'     For i = 0 To adors.Fields.Count - 1
'         Sheet14.Cells(1, i + 1) = adors.Fields(i).Name
'     Next i
     
     
     Sheet7.Range("a1048576").End(3).Offset(1, 0).CopyFromRecordset adors
     adors.Close
     Set adors = Nothing
    'Sheet14.Columns("A:A").NumberFormat = "@"  61721
    Application.ScreenUpdating = True
    
End Sub

Sub turnsDtoH()   'for Wanek3_DC_OUT

't = Timer
Application.ScreenUpdating = False
Sheet5.Activate
Dim arr, row&
row = Sheet5.Range("ab1048576").End(3).row
Sheet5.Range("a2:h1048576").ClearContents
arr = Sheet5.Range("a1").CurrentRegion

Dim x
For x = 2 To UBound(arr)
    arr(x, 7) = CDate(arr(x, 28))
    arr(x, 8) = Application.WeekNum(arr(x, 7))
    arr(x, 6) = arr(x, 20) & "+" & arr(x, 9)
    arr(x, 5) = arr(x, 14) & "-" & arr(x, 26) & "-" & arr(x, 27)
    arr(x, 4) = arr(x, 22) * 1
Next x
    
With Sheet5.Range("a1").CurrentRegion

    .Value = arr
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    Columns("a:h").AutoFit
End With
    Sheet5.Columns("g:g").NumberFormat = "mm-dd-yyyy"

    Erase arr
    
'  Dim r%, i%
'  Dim arr, brr
'  With Sheet5
'    r = .Cells(.Rows.Count, 1).End(xlUp).row
'    arr = .Range("a2:aj" & r)
'    For i = 1 To UBound(arr)
'      arr(i, 7) = CDate(arr(i, 28))
'    Next
'    .Range("a2:aj" & r) = arr
'  End With
   
    'MsgBox "udpated~ " & Timer
Application.ScreenUpdating = True
End Sub


Sub UPfind()  'ÒýÓÃUnite Price   'for Wanek3_DC_OUT

't = Timer
Application.ScreenUpdating = False
Dim d As Object, arr, brr, i&
row = Sheet5.Range("i1048576").End(3).row

Set d = CreateObject("scripting.dictionary")
    d.CompareMode = vbTextCompare '²»Çø·Ö×ÖÄ¸´óÐ¡Ð´
    
arr = Sheet7.Range("a1").CurrentRegion  'Êý¾ÝÔ´×°ÈëÊý×éarr
brr = Sheet5.Range("a1").CurrentRegion  '²éÑ¯ÇøÓòÊý¾Ý×°ÈëÊý×ébrr

For i = 1 To UBound(arr)            '±éÀúÊý×éarr
    d(arr(i, 1)) = arr(i, 2)        '½«item + up ×÷Îªkey£¬×°Èë×Öµä
Next

For i = 2 To UBound(brr)            '±êÌâÐÐ²»ÓÃ²éÑ¯£¬ËùÒÔ´ÓµÚ¶þÐÐ¿ªÊ¼±éÀú²éÑ¯ÊýÖµbrr
    If d.exists(brr(i, 9)) Then brr(i, 3) = d(brr(i, 9)) Else brr(i, 3) = 200
    'Èç¹û×ÖµäÖÐ´æÔÚitem, '¸ù¾Ýitem ´Ó×ÖµäÖÐÈ¡UPÖµ
    brr(i, 2) = brr(i, 3) * brr(i, 4)
    
Next i

 'Sheet5.[a1].Resize(UBound(brr), UBound(brr, 2)) = brr   '(brr,2)ÊÇÖ¸Êý×éµÄ¶þÎ³ÏÂ±ê
With Sheet5.Range("a1").CurrentRegion

    .Value = brr
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter

End With

    Set d = Nothing        'ÊÍ·Å×Öµä
    Erase arr
    Erase brr
    
Application.ScreenUpdating = True
    
End Sub


Sub Sitevlookup_A()  ' Wanek3_DC_OUT column A for site


't = Timer
Application.ScreenUpdating = False
Dim d As Object, arr, brr, i&
row = Sheet5.Range("i1048576").End(3).row

Set d = CreateObject("scripting.dictionary")
    d.CompareMode = vbTextCompare '²»Çø·Ö×ÖÄ¸´óÐ¡Ð´
    
arr = Sheet8.Range("a1").CurrentRegion  'Êý¾ÝÔ´×°ÈëÊý×éarr
brr = Sheet5.Range("a1").CurrentRegion  '²éÑ¯ÇøÓòÊý¾Ý×°ÈëÊý×ébrr

For i = 1 To UBound(arr)            '±éÀúÊý×éarr
    d(arr(i, 1)) = arr(i, 2)        '½«item + up ×÷Îªkey£¬×°Èë×Öµä
Next

For i = 2 To UBound(brr)            '±êÌâÐÐ²»ÓÃ²éÑ¯£¬ËùÒÔ´ÓµÚ¶þÐÐ¿ªÊ¼±éÀú²éÑ¯ÊýÖµbrr
    If d.exists(brr(i, 5)) Then brr(i, 1) = d(brr(i, 5)) Else brr(i, 1) = "VIEW"
    'Èç¹û×ÖµäÖÐ´æÔÚitem, '¸ù¾Ýitem ´Ó×ÖµäÖÐÈ¡UPÖµ

    
Next i

 'Sheet5.[a1].Resize(UBound(brr), UBound(brr, 2)) = brr   '(brr,2)ÊÇÖ¸Êý×éµÄ¶þÎ³ÏÂ±ê
With Sheet5.Range("a1").CurrentRegion

    .Value = brr
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter

    'MsgBox "udpated~ " & Timer - t
    
End With
    Set d = Nothing        'ÊÍ·Å×Öµä
    Erase arr
    Erase brr
Application.ScreenUpdating = True
End Sub



Sub NoSiteInColumnASorting()   'update the site sheet

Application.ScreenUpdating = False
Sheet8.Activate
Set cnn = CreateObject("adodb.connection")
cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=Excel 12.0;Data Source=" & ThisWorkbook.FullName
sql = "select Distinct [WH+EMP+SUP] " _
    & "from [Wanek3_DC_BW_OUT$] " _
    & "where [SITE]=""VIEW"" " _
    & "Order by [WH+EMP+SUP]"

Sheet8.[a33653].End(3).Offset(1, 0).CopyFromRecordset cnn.execute(sql)
cnn.Close
Set cnn = Nothing
Sheet8.Select
Application.ScreenUpdating = True
MsgBox "updated"
End Sub



Sub WANEK3_IN()  'for Wanek3 in column C and D
'
't = Timer
Application.ScreenUpdating = False
Sheet1.Activate
Sheet1.Range("a2:d1048576").ClearContents

Dim arr, brr(1 To 1048576, 1 To 2), row&
arr = Sheet1.Range("a1").CurrentRegion
row = Sheet1.Range("e1048576").End(3).row

Dim x&
For x = 2 To UBound(arr)
    arr(x, 3) = arr(x, 18) * 1
    arr(x, 4) = CDate(arr(x, 24))

    Next x
    
    With Sheet1.Range("a1").CurrentRegion

    .Value = arr
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
     Columns("d:d").NumberFormat = "mm-dd-yyyy"
    End With
    Erase arr

    
Application.ScreenUpdating = True
    'MsgBox "udpated~ " & Timer - t
End Sub
    

Sub UPfind_in()  'ÒýÓÃUnite Price  'for Wanek3 in column B

't = Timer
Application.ScreenUpdating = False
Sheet1.Activate
Dim d As Object, arr, brr, i&
row = Sheet1.Range("e1048576").End(3).row


Set d = CreateObject("scripting.dictionary")
    d.CompareMode = vbTextCompare '²»Çø·Ö×ÖÄ¸´óÐ¡Ð´
    
arr = Sheet7.Range("a1").CurrentRegion  'Êý¾ÝÔ´×°ÈëÊý×éarr
brr = Sheet1.Range("a1").CurrentRegion  '²éÑ¯ÇøÓòÊý¾Ý×°ÈëÊý×ébrr

For i = 1 To UBound(arr)            '±éÀúÊý×éarr
    d(arr(i, 1)) = arr(i, 2)        '½«item + up ×÷Îªkey£¬×°Èë×Öµä
Next

For i = 2 To UBound(brr)            '±êÌâÐÐ²»ÓÃ²éÑ¯£¬ËùÒÔ´ÓµÚ¶þÐÐ¿ªÊ¼±éÀú²éÑ¯ÊýÖµbrr
    If d.exists(brr(i, 5)) Then brr(i, 2) = d(brr(i, 5)) Else brr(i, 2) = 200
    'Èç¹û×ÖµäÖÐ´æÔÚitem, '¸ù¾Ýitem ´Ó×ÖµäÖÐÈ¡UPÖµ
    brr(i, 1) = brr(i, 2) * brr(i, 3)
    
Next i

 'Sheet5.[a1].Resize(UBound(brr), UBound(brr, 2)) = brr   '(brr,2)ÊÇÖ¸Êý×éµÄ¶þÎ³ÏÂ±ê
With Sheet1.Range("a1").CurrentRegion

    .Value = brr
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter

End With

    Set d = Nothing        'ÊÍ·Å×Öµä
    Erase arr
    Erase brr
Application.ScreenUpdating = True
    'MsgBox "udpated~ " & Timer - t
End Sub


Sub WANEK3_202_OUT()  'for WANEK3_202_OUT column C and D

't = Timer
Application.ScreenUpdating = False

Sheet2.Activate
Sheet2.Range("a2:d1048576").ClearContents

Dim arr, brr(1 To 1048576, 1 To 2), row&
arr = Sheet2.Range("a1").CurrentRegion
row = Sheet2.Range("e1048576").End(3).row


Dim x&
For x = 2 To UBound(arr)
    arr(x, 3) = arr(x, 18) * 1
    arr(x, 4) = CDate(arr(x, 24))

    Next x
    
    With Sheet2.Range("a1").CurrentRegion

    .Value = arr
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    Range("d2:d" & row).NumberFormat = "mm-dd-yyyy"
    
    End With
    Erase arr
   
 Application.ScreenUpdating = True
      
    'MsgBox "udpated~ " & Timer - t
End Sub
    

Sub UPfind_in_WANEK3_202_OUT()  'ÒýÓÃUnite Price  'for Wanek3 in column B

't = Timer
Application.ScreenUpdating = False
Sheet2.Activate
Dim d As Object, arr, brr, i&
row = Sheet2.Range("e1048576").End(3).row


Set d = CreateObject("scripting.dictionary")
    d.CompareMode = vbTextCompare '²»Çø·Ö×ÖÄ¸´óÐ¡Ð´
    
arr = Sheet7.Range("a1").CurrentRegion  'Êý¾ÝÔ´×°ÈëÊý×éarr
brr = Sheet2.Range("a1").CurrentRegion  '²éÑ¯ÇøÓòÊý¾Ý×°ÈëÊý×ébrr

For i = 1 To UBound(arr)            '±éÀúÊý×éarr
    d(arr(i, 1)) = arr(i, 2)        '½«item + up ×÷Îªkey£¬×°Èë×Öµä
Next

For i = 2 To UBound(brr)            '±êÌâÐÐ²»ÓÃ²éÑ¯£¬ËùÒÔ´ÓµÚ¶þÐÐ¿ªÊ¼±éÀú²éÑ¯ÊýÖµbrr
    If d.exists(brr(i, 5)) Then brr(i, 2) = d(brr(i, 5)) Else brr(i, 2) = 200
    'Èç¹û×ÖµäÖÐ´æÔÚitem, '¸ù¾Ýitem ´Ó×ÖµäÖÐÈ¡UPÖµ
    brr(i, 1) = brr(i, 2) * brr(i, 3)
Next i

 'Sheet5.[a1].Resize(UBound(brr), UBound(brr, 2)) = brr   '(brr,2)ÊÇÖ¸Êý×éµÄ¶þÎ³ÏÂ±ê
With Sheet2.Range("a1").CurrentRegion

    .Value = brr
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter

End With
    Set d = Nothing        'ÊÍ·Å×Öµä
    Erase arr
    Erase brr
   ' MsgBox "udpated~ " & Timer - t
Application.ScreenUpdating = True
End Sub


Sub DC_IN()  'for DC_IN column C and D

't = Timer
Application.ScreenUpdating = False
Sheet4.Activate
Sheet4.Range("a2:d1048576").ClearContents

Dim arr, brr(1 To 1048576, 1 To 2), row&
arr = Sheet4.Range("a1").CurrentRegion
row = Sheet4.Range("e1048576").End(3).row


Dim x&
For x = 2 To UBound(arr)
    arr(x, 3) = arr(x, 18) * 1
    arr(x, 4) = CDate(arr(x, 24))

    Next x
    
    With Sheet4.Range("a1").CurrentRegion

    .Value = arr
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    Range("d2:d" & row).NumberFormat = "mm-dd-yyyy"
    End With
    Erase arr
   
Application.ScreenUpdating = True
    'MsgBox "udpated~ " & Timer - t
End Sub
    

Sub UPfind_in_DC_IN()  'ÒýÓÃUnite Price  'for DC_IN column B

't = Timer
Application.ScreenUpdating = False
Sheet4.Activate
Dim d As Object, arr, brr, i&
row = Sheet4.Range("e1048576").End(3).row


Set d = CreateObject("scripting.dictionary")
    d.CompareMode = vbTextCompare '²»Çø·Ö×ÖÄ¸´óÐ¡Ð´
    
arr = Sheet7.Range("a1").CurrentRegion  'Êý¾ÝÔ´×°ÈëÊý×éarr
brr = Sheet4.Range("a1").CurrentRegion  '²éÑ¯ÇøÓòÊý¾Ý×°ÈëÊý×ébrr

For i = 1 To UBound(arr)            '±éÀúÊý×éarr
    d(arr(i, 1)) = arr(i, 2)        '½«item + up ×÷Îªkey£¬×°Èë×Öµä
Next

For i = 2 To UBound(brr)            '±êÌâÐÐ²»ÓÃ²éÑ¯£¬ËùÒÔ´ÓµÚ¶þÐÐ¿ªÊ¼±éÀú²éÑ¯ÊýÖµbrr
    If d.exists(brr(i, 5)) Then brr(i, 2) = d(brr(i, 5)) Else brr(i, 2) = 200
    'Èç¹û×ÖµäÖÐ´æÔÚitem, '¸ù¾Ýitem ´Ó×ÖµäÖÐÈ¡UPÖµ
    brr(i, 1) = brr(i, 2) * brr(i, 3)
Next i

 'Sheet5.[a1].Resize(UBound(brr), UBound(brr, 2)) = brr   '(brr,2)ÊÇÖ¸Êý×éµÄ¶þÎ³ÏÂ±ê
With Sheet4.Range("a1").CurrentRegion

    .Value = brr
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter

End With
    Set d = Nothing        'ÊÍ·Å×Öµä
    Erase arr

Application.ScreenUpdating = True
   ' MsgBox "udpated~ " & Timer - t
End Sub


Sub STO_UP_BCD()  'for STO sheet B~D

't = Timer
Application.ScreenUpdating = False
Sheet3.Activate
Sheet3.Range("a2:d1048576").ClearContents

Dim d As Object, arr, brr, i&
row = Sheet3.Range("f1048576").End(3).row


Set d = CreateObject("scripting.dictionary")
    d.CompareMode = vbTextCompare '²»Çø·Ö×ÖÄ¸´óÐ¡Ð´
    
arr = Sheet7.Range("a1").CurrentRegion  'Êý¾ÝÔ´×°ÈëÊý×éarr
brr = Sheet3.Range("a1").CurrentRegion  '²éÑ¯ÇøÓòÊý¾Ý×°ÈëÊý×ébrr

For i = 1 To UBound(arr)            '±éÀúÊý×éarr
    d(arr(i, 1)) = arr(i, 2)        '½«item + up ×÷Îªkey£¬×°Èë×Öµä
Next

For i = 2 To UBound(brr)            '±êÌâÐÐ²»ÓÃ²éÑ¯£¬ËùÒÔ´ÓµÚ¶þÐÐ¿ªÊ¼±éÀú²éÑ¯ÊýÖµbrr
    If d.exists(brr(i, 6)) Then brr(i, 4) = d(brr(i, 6)) Else brr(i, 4) = 200
    'Èç¹û×ÖµäÖÐ´æÔÚitem, '¸ù¾Ýitem ´Ó×ÖµäÖÐÈ¡UPÖµ
    brr(i, 2) = brr(i, 7) * 1
    brr(i, 3) = brr(i, 2) * brr(i, 4)
    
Next i

With Sheet3.Range("a1").CurrentRegion

    .Value = brr

    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With
    Set d = Nothing        'ÊÍ·Å×Öµä
    Erase arr
    Erase brr
Application.ScreenUpdating = True
   ' MsgBox "udpated~ " & Timer - t
End Sub

Sub STO_Site_A()  'for STO sheet column A

Sheet3.Activate
Application.ScreenUpdating = False
Dim d As Object, arr, brr, i&
row = Sheet3.Range("f1048576").End(3).row


Set d = CreateObject("scripting.dictionary")
    d.CompareMode = vbTextCompare '²»Çø·Ö×ÖÄ¸´óÐ¡Ð´
    
arr = Sheet11.Range("a1").CurrentRegion  'Êý¾ÝÔ´×°ÈëÊý×éarr
brr = Sheet3.Range("a1").CurrentRegion  '²éÑ¯ÇøÓòÊý¾Ý×°ÈëÊý×ébrr

For i = 1 To UBound(arr)            '±éÀúÊý×éarr
    d(arr(i, 1)) = arr(i, 2)        '½«item + up ×÷Îªkey£¬×°Èë×Öµä
    Next

For i = 2 To UBound(brr)            '±êÌâÐÐ²»ÓÃ²éÑ¯£¬ËùÒÔ´ÓµÚ¶þÐÐ¿ªÊ¼±éÀú²éÑ¯ÊýÖµbrr
    If d.exists(brr(i, 11)) Then brr(i, 1) = d(brr(i, 11)) Else brr(i, 1) = "VIEW"
    Next i

With Sheet3.Range("a1").CurrentRegion

    .Value = brr
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    Columns("a:d").AutoFit
End With
    Set d = Nothing        'ÊÍ·Å×Öµä
    Erase arr
    Erase brr
Application.ScreenUpdating = True
 
End Sub

Sub STONoSiteInColumnASorting()   'update STO A column site sheet

Sheet3.Activate
Application.ScreenUpdating = False
Set cnn = CreateObject("adodb.connection")
cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=Excel 12.0;Data Source=" & ThisWorkbook.FullName
sql = "select Distinct [Site], [Location Id] " _
    & "from [STO$] " _
    & "where [SITE]=""VIEW"" " _
    & "Order by [Site]"

Sheet11.[a33653].End(3).Offset(1, 0).CopyFromRecordset cnn.execute(sql)
cnn.Close
Set cnn = Nothing

Sheet11.Select
Application.ScreenUpdating = True
'MsgBox "updated"

End Sub


Sub allsheetscdate()   'for Wanek3_DC_OUT

't = Timer
Application.ScreenUpdating = False
Sheet5.Activate
Dim arr, row&

row = Sheet5.Range("i1048576").End(3).row
arr = Sheet5.Range("g2:g" & row)

Dim x
For x = 2 To UBound(arr)
k = k + 1
    arr(k, 1) = CDate(arr(x, 1))
    Next
   
With Sheet5.Range("g2:g" & row)
    .NumberFormat = "mm-dd-yyyy"
    .Value = arr
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With
    Erase arr

Application.ScreenUpdating = True
'MsgBox "udpated~ " & Format(Timer - t, "0.00") & "s"
End Sub



Sub allsheetscdate2()   'for WANEK3_IN

't = Timer
Application.ScreenUpdating = False
Sheet1.Activate
Dim arr, row&

row = Sheet1.Range("e1048576").End(3).row
arr = Sheet1.Range("d2:d" & row)

Dim x
For x = 2 To UBound(arr)
k = k + 1
    arr(k, 1) = CDate(arr(x, 1))
    Next
   
With Sheet1.Range("d2:d" & row)
    .NumberFormat = "mm-dd-yyyy"
    .Value = arr
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With
    Erase arr
 
Application.ScreenUpdating = True
    'MsgBox "udpated~ " & Format(Timer - t, "0.00") & "s"
End Sub

Sub allsheetscdate3()   'for WANEK3_202_OUT

't = Timer
Application.ScreenUpdating = False
Sheet2.Activate
Dim arr, row&

row = Sheet2.Range("e1048576").End(3).row
arr = Sheet2.Range("d2:d" & row)

Dim x
For x = 2 To UBound(arr)
k = k + 1
    arr(k, 1) = CDate(arr(x, 1))
    Next
   
With Sheet2.Range("d2:d" & row)
    .NumberFormat = "mm-dd-yyyy"
    .Value = arr
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With
    Erase arr

Application.ScreenUpdating = True
    'MsgBox "udpated~ " & Format(Timer - t, "0.00") & "s"
End Sub

Sub allsheetscdate4()   'for DC_IN

't = Timer
Application.ScreenUpdating = False
Sheet4.Activate
Dim arr, row&

row = Sheet4.Range("e1048576").End(3).row
arr = Sheet4.Range("d2:d" & row)

Dim x
For x = 2 To UBound(arr)
k = k + 1
    arr(k, 1) = CDate(arr(x, 1))
    Next
   
With Sheet4.Range("d2:d" & row)
    .NumberFormat = "mm-dd-yyyy"
    .Value = arr
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With
    Erase arr

Application.ScreenUpdating = True

    'MsgBox "udpated~ " & Format(Timer - t, "0.00") & "s"
End Sub

Sub allsheetscdate5()   'for BW_202_OUT

't = Timer
Application.ScreenUpdating = False
Sheet6.Activate
Dim arr, row&

row = Sheet6.Range("e1048576").End(3).row
arr = Sheet6.Range("d2:d" & row)

Dim x
For x = 2 To UBound(arr)
k = k + 1
    arr(k, 1) = CDate(arr(x, 1))
    Next
   
With Sheet6.Range("d2:d" & row)
    .NumberFormat = "mm-dd-yyyy"
    .Value = arr
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With
    Erase arr

Application.ScreenUpdating = True

'MsgBox "udpated~ " & Format(Timer - t, "0.00") & "s"
End Sub








