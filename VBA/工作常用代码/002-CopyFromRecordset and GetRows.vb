Sub CubeAndDescriptionQueryV04()

t = Timer
Sheet1.Select
Sheet1.Range("b2:c1048576").Clear
Application.ScreenUpdating = False
Dim i%, nrow&, item As String, arr(), brr(), crr(), items As String, k&

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
    adors.Open cmdtxt, Db, 3, 3
    
     arr = adors.GetRows()
     arr = Application.Transpose(arr)
     adors.Close
     Set adors = Nothing

    'xlookup description and unit of measure to HJ sheet
    Dim j&, lrow&, d As Object
    Set d = CreateObject("scripting.dictionary")
    
    brr = Sheet1.Range("a1:c" & Range("a1048576").End(3).Row)
    Set d = CreateObject("scripting.dictionary")
    d.CompareMode = vbTextCompare
    For j = 1 To UBound(arr)
        d(arr(j, 2)) = arr(j, 6) & ";" & arr(j, 4)
    Next

    For j = 2 To UBound(brr)
        If d.exists(brr(j, 1)) Then brr(j, 2) = Split(d(brr(j, 1)), ";")(0) Else brr(j, 2) = "item number is not correct, please check"
        If d.exists(brr(j, 1)) Then brr(j, 3) = Split(d(brr(j, 1)), ";")(1) Else brr(j, 3) = "item number is not correct, please check"
    Next
    Sheet1.Columns("a").NumberFormat = "@"
    Sheet1.Columns("c").NumberFormat = "@"
    Sheet1.Range("a1:c" & Range("a1048576").End(3).Row).Value = brr
    
    Erase brr
    Erase arr
    Set d = Nothing
    ThisWorkbook.Save
    Sheet1.Select
    Application.ScreenUpdating = True
    MsgBox "Query Successful in " & Format(Timer - t, "0.00" & "s") & "!"
End Sub








