Sub CubeAndDescriptionQueryV04()

t = Timer
Sheet1.Select
Sheet1.Range("b2:c1048576").Clear
Application.ScreenUpdating = False
    Dim i%, nrow&, item As String, arr(), brr(), crr(), items As String, k&, skus As String, d1 As Object, aRes
    'combination all items as string
    For i = 1 To 100
        If Range("a" & i) <> "" Then
            GoTo 100
        Else
            MsgBox "Column A no item number"
            Exit Sub
        End If
    Next
100

Sheet1.Activate
arr = Sheet1.Range("a2:a" & Range("a1048576").End(3).Row)

    Set d1 = CreateObject("scripting.dictionary")
    d1.CompareMode = vbTextCompare
    For k = 1 To UBound(arr)
       If Not d1.exists(arr(k, 1)) Then
        d1(arr(k, 1)) = ""
        End If
    Next
    aRes = d1.keys
    
    For i = 0 To UBound(aRes)
        If i < UBound(aRes) Then
            item = aRes(i)
            skus = "'" & item & "'" & "," & skus
        Else
            item = aRes(i)
            skus = skus & "'" & item & "'"
        End If
    Next

'PULL ITEMS
    Dim cmdtxt As String
    Dim adors As New Recordset
    Set Db = New Connection
    Db.CursorLocation = adUseClient
    If Db.State = 1 Then Db.Close
    U = Sheet4.Range("a1").Value
    P = Sheet4.Range("a2").Value
    Db.Open "Provider =IBMDASQL.DataSource.1" & _
     ";Catalog Library List=JDETSTDTA" & _
     ";Persist Security Info=True" & _
     ";Force Translate=0" & _
     ";Data Source = WFVNPROD " & _
     ";User ID =" & U & "" & _
     ";Password =" & P
     Set adors = New Recordset
     If adors.State = 1 Then adors.Close
    cmdtxt = "SELECT t1.STID,T1.ITNBR,T1.ITCLS,T1.ITDSC,T1.UNMSR,T1.B2Z95S " & _
         "FROM AMFLIBW.ITMRVA T1 " & _
         "WHERE T1.STID IN ('35') "
      
    If skus <> "" Then
          cmdtxt = cmdtxt & " AND T1.ITNBR in " & "(" & skus & ")"
    End If
    cmdtxt = cmdtxt
    
    adors.Open cmdtxt, Db, 3, 3
    adors.MoveFirst
    crr = adors.GetRows()
    'arr = Application.Transpose(crr)
    adors.Close
    Set adors = Nothing
    
    'xlookup description and unit of measure to HJ sheet
    Dim j&, lrow&, d As Object
    Set d = CreateObject("scripting.dictionary")
    
    brr = Sheet1.Range("a1:c" & Range("a1048576").End(3).Row)
    Set d = CreateObject("scripting.dictionary")
    d.CompareMode = vbTextCompare
    For j = 0 To UBound(crr, 2)
        d(crr(1, j)) = crr(5, j) & ";" & crr(3, j)
    Next
    
    For j = 2 To UBound(brr)
        If d.exists(brr(j, 1)) Then brr(j, 2) = Split(d(brr(j, 1)), ";")(0) Else brr(j, 2) = "item number is not correct, please check"
        If d.exists(brr(j, 1)) Then brr(j, 3) = Split(d(brr(j, 1)), ";")(1) Else brr(j, 3) = "item number is not correct, please check"
    Next
    
    Sheet1.Columns("a").NumberFormat = "@"
    Sheet1.Columns("c").NumberFormat = "@"
    Sheet1.Range("a1:c" & Range("a1048576").End(3).Row).Value = brr
    Erase arr
    Erase brr
    Erase crr
    Erase aRes
    Set d = Nothing
    Set d1 = Nothing
    ThisWorkbook.Save
    Sheet1.Select
    Application.ScreenUpdating = True
    MsgBox "Query Successful in " & Format(Timer - t, "0.00" & "s") & "!"
End Sub








