'WANEK3 Invoice description and unit of measure, created by Jimshen on Dec.23.2021

Sub string_items()   'V03

t = Timer
Application.ScreenUpdating = False
Dim i%, nrow&, item As String, arr(), brr(), crr(), items As String, k&

'combination all items as string
Sheet2.Select
For i = 1 To 100
    If Range("c" & i) <> "" Then
        GoTo 100
    Else
        MsgBox "Column C no item number"
        Exit Sub
    End If
Next
100
nrow = Sheet2.Range("c1048576").End(3).Row
arr = Range("a1:c" & Range("c1048576").End(3).Row)

    For i = 1 To nrow
       If i < nrow Then
            arr(i, 3) = Range("c" & i)
            arr(i, 3) = "'" & arr(i, 3) & "'" & ","
            item = arr(i, 3)
            items = items + item
        ElseIf i = nrow Then
            arr(i, 3) = Range("c" & i)
            arr(i, 3) = "'" & arr(i, 3) & "'"
            item = arr(i, 3)
            items = items + item
        End If
    Next

' pull AS400 data
sku = Worksheets("HJ").Range("c1").Text
Dim cmdtxt As String
Dim adors As New Recordset
Dim sh As Worksheet
Dim adoCN As Object
Dim strSQL As String

    Set Db = New Connection
    Db.CursorLocation = adUseClient
    If Db.State = 1 Then Db.Close
    U = Sheet6.Range("a1").Value
    P = Sheet6.Range("a2").Value
    Db.Open "Provider =IBMDASQL.DataSource.1" & _
     ";Catalog Library List=JDETSTDTA" & _
     ";Persist Security Info=True" & _
     ";Force Translate=0" & _
     ";Data Source = WFVNPROD " & _
     ";User ID =" & U & "" & _
     ";Password =" & P
     Set adors = New Recordset
     If adors.State = 1 Then adors.Close
    cmdtxt = "SELECT t1.STID,T1.ITNBR,T1.ITCLS,T1.ITDSC,T1.UNMSR " & _
         "FROM AMFLIBW.ITMRVA T1 " & _
         "WHERE T1.STID IN ('35','33','31') "
         
    '    "WHERE T1.ITCLS LIKE 'Z%' AND T1.ITCLS NOT LIKE '%K' AND T1.STID IN ('35') "
     
    If sku <> "" Then
          cmdtxt = cmdtxt & " AND T1.ITNBR in " & "(" & items & ")"
    End If
    cmdtxt = cmdtxt
    adors.Open cmdtxt, Db, 3, 3
    Sheet3.Select
    Sheet3.Cells.Clear
    Sheet3.Columns("A:d").NumberFormat = "@"
     For i = 0 To adors.Fields.Count - 1
         Sheet3.Cells(1, i + 1) = adors.Fields(i).Name
     Next i
     Sheet3.Range("a1048576").End(3).Offset(1, 0).CopyFromRecordset adors
     adors.Close
     Erase arr
     Set adors = Nothing

    'xlookup description and unit of measure to HJ sheet
    Dim j&, lrow&, d As Object
    Set d = CreateObject("scripting.dictionary")
    Sheet2.Activate
    Sheet2.Range("a1:b10000").ClearContents
    Sheet2.Range("a1:b10000").Interior.Pattern = xlNone
    lrow = Sheet3.Range("a1048576").End(3).Row
    arr = Sheet2.Range("a1:c" & Range("c1048576").End(3).Row)
    brr = Sheet3.Range("a1:e" & lrow)
    Set d = CreateObject("scripting.dictionary")
    d.CompareMode = vbTextCompare
    For j = 1 To UBound(brr)
        d(brr(j, 2)) = brr(j, 5) & "/" & brr(j, 4)
    Next
    Sheet2.Activate
    For j = 1 To Range("c1048576").End(3).Row
        If d.exists(arr(j, 3)) Then Range("a" & j) = Split(d(arr(j, 3)), "/")(1) Else Range("a" & j) = "item number is not correct, please check"
        If d.exists(arr(j, 3)) Then Range("b" & j) = Split(d(arr(j, 3)), "/")(0) Else Range("b" & j) = "item number is not correct, please check"
    
    Next
    For k = 1 To Range("c1048576").End(3).Row
        If Cells(k, 1) = "item number is not correct, please check" Then
            Cells(k, 1).Interior.ColorIndex = 3
            Cells(k, 2).Interior.ColorIndex = 3
        Else
            Cells(k, 1).Interior.ColorIndex = 43
            Cells(k, 2).Interior.ColorIndex = 43
        End If
    Next
    Erase brr
    Set d = Nothing
    ThisWorkbook.Save
    Application.ScreenUpdating = True
    MsgBox "Query Successful in " & Format(Timer - t, "0.00" & "s") & "!"
End Sub


