'Pull TripDetailInformation
Sub load_TripDetails()
    
    Application.ScreenUpdating = False
    Dim i As Long, t As String
    Dim adors As New Recordset
    Sheet1.Activate
    Sheet1.Cells.Clear
    
    
    Set Db = New Connection
    Db.CursorLocation = adUseClient
    If Db.State = 1 Then Db.Close
    
    U = Sheet4.Range("a1").Value
    P = Sheet4.Range("a2").Value
    X = Sheet4.Range("b4").Value
    Y = Sheet4.Range("C4").Value
    Z = Sheet4.Range("D4").Value
    XX = Sheet4.Range("E4").Value
    YY = Sheet4.Range("F4").Value
    ZZ = Sheet4.Range("G4").Value
    
    
    
    't = Application.InputBox(Prompt:="请输入trip号码: ", Type:=3)
    Db.Open "Provider =IBMDASQL.DataSource.1" & _
            ";Catalog Library List=JIMTDTA" & _
            ";Persist Security Info=True" & _
            ";Force Translate=0" & _
            ";Data Source = WVFHA" & _
            ";User ID =" & U & "" & _
            ";Password =" & P
    
    Set adors = New Recordset
    If adors.State = 1 Then adors.Close
    
     cmdtxt = "Select t2.BDITM#,t2.BDITMD,t2.BDICLS,sum(t2.BDITQT) as Qty, sum(t2.BDITCT) as Cubes, sum(t2.BDITWT) as Weight " & _
              "from (Select BDTRP#,BDITM#,BDITMD,BDICLS,BDITQT,BDITCT,BDITWT From DISTLIBQ.BTTRIPD t1 Where BDTRP# IN (" & X & "," & Y & "," & Z & "," & XX & "," & YY & "," & ZZ & ") order by BDTRP#,bditm#) as t2 " & _
              "group by t2.BDITM#,t2.BDITMD,t2.BDICLS "

     
    adors.Open cmdtxt, Db, 3, 3
    For i = 0 To adors.Fields.Count - 1
        Sheet1.Cells(1, i + 1) = adors.Fields(i).Name
    Next i
    
    Sheet1.Range("a2").CopyFromRecordset adors
    adors.Close
    Set adors = Nothing
    
    Sheet1.Columns("A:C").NumberFormat = "@"

    Application.ScreenUpdating = True
End Sub


'Pull_AS_STO_REPORT

Sub Pull_AS_STO() '  ADO读取其他工作簿
    Dim wb As Workbook
    Dim arr, brr(), i&, j&, k&, nrow&, crr()
    
    't = Timer
    Application.ScreenUpdating = False
    
    Sheet2.Select
    Cells.Clear
    Range("h1").Value = "Data collected at:" & Format(Now(), "hhmm,mm-dd-yyyy")
    
    Set wb = GetObject("C:\Users\jishen\Downloads\ASSTO.xlsx") '打开工作簿
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False
    ReDim brr(1 To UBound(arr), 1 To UBound(arr, 2))
    
    For i = 1 To UBound(arr)
            brr(i, 1) = arr(i, 2)
            brr(i, 2) = arr(i, 3)
            brr(i, 3) = arr(i, 7)
        
    
'        For j = 1 To UBound(arr, 2)
'            brr(i, j) = arr(i, j)
'        Next
    Next
    
    Columns("a:a").NumberFormat = "@"
    Columns("c:d").NumberFormat = "@"
    Sheet2.Range("a1").Resize(UBound(arr), UBound(arr, 2)) = brr
    
    Columns("a:f").EntireColumn.AutoFit
 
    
    Application.ScreenUpdating = True
    'MsgBox Format(Timer - t, "0.00" & "s")
    
End Sub


Sub add_LOC_QTY()   'sheet2

    Application.ScreenUpdating = False
    Dim i&, arr(), nrow&
    Sheet2.Range("d1:e1") = Array("Loc+Qty", "Building")
    nrow = Sheet2.Range("a1048576").End(3).Row
    arr = Sheet2.Range("a1:e" & nrow).CurrentRegion
    
    With Sheet2
        For i = 2 To UBound(arr)
            arr(i, 4) = arr(i, 3) & "*" & arr(i, 2) & "pcs"
            If Mid(arr(i, 3), 1, 2) = "K6" Then
                arr(i, 5) = "B6"
            ElseIf Mid(arr(i, 3), 1, 2) = "K5" Then
                arr(i, 5) = "B5"
            ElseIf Mid(arr(i, 3), 1, 2) = "RS" Then
                arr(i, 5) = "B5"
            ElseIf Mid(arr(i, 3), 1, 2) = "ST" Then
                arr(i, 5) = "B6"
            ElseIf Mid(arr(i, 3), 1, 2) = "WR" Then
                arr(i, 5) = "B6"
            ElseIf Mid(arr(i, 3), 1, 2) = "FX" Then
                arr(i, 5) = "B5"
            Else
                arr(i, 5) = "Other"
            End If
        Next
    
        .Range("a1").Resize(UBound(arr), UBound(arr, 2)).Value = arr
        .Columns("a").NumberFormat = "@"
        .Columns("a:d").EntireColumn.AutoFit
        
    End With
    
    Application.ScreenUpdating = True

End Sub


Sub DATA_Columns_Ghij()

    Application.ScreenUpdating = False
    Dim d As Object
    Dim i&, arr(), brr()
    Set d1 = CreateObject("scripting.dictionary")
    Set d2 = CreateObject("scripting.dictionary")
    Set d3 = CreateObject("scripting.dictionary")
    Set d4 = CreateObject("scripting.dictionary")
    
    Sheet1.Range("G1:J1").Value = Array("B5 Qty", "B6 Qty", "需补货数量", "Locations")
    arr = Sheet2.Range("a1").CurrentRegion '需要汇总的来源表STO
    
    For i = 2 To UBound(arr)
        If arr(i, 5) = "B5" Then
            d1(arr(i, 1)) = d1(arr(i, 1)) + arr(i, 2) '以item number为key,将Qty累加到字典1
        ElseIf arr(i, 5) = "B6" Then
            d2(arr(i, 1)) = d2(arr(i, 1)) + arr(i, 2) '以item number为key,将Qty累加到字典2
        'd(arr(i, 2)) = d(arr(i, 2)) & " " & arr(i, 7) & "*" & arr(i, 3) & "pcs" '以item number为key,将location与qty连接
        Else
            d3(arr(i, 1)) = d2(arr(i, 1)) + arr(i, 2) '以item number为key,将Qty累加到字典2
        End If
        d4(arr(i, 1)) = d4(arr(i, 1)) & "," & arr(i, 4) '以item number为key,将location*qty 进行连接
        
         
    
    Next i
    
    
    brr = Sheet1.Range("a1").CurrentRegion
    For i = 2 To UBound(brr)
        If d1.exists(brr(i, 1)) Then
            brr(i, 7) = d1(brr(i, 1))
        Else
                brr(i, 7) = 0          '以item number来查询字典，将B5 Qty查询到目的地来
        End If
        If d2.exists(brr(i, 1)) Then
            brr(i, 8) = d2(brr(i, 1))
        Else
                brr(i, 8) = 0          '以item number来查询字典，将B6 Qty查询到目的地来
        End If
        If d4.exists(brr(i, 1)) Then
            brr(i, 10) = d4(brr(i, 1))
        Else
                brr(i, 10) = ""          '以item number来查询字典，将Other Qty查询到目的地来
        End If
        
        If brr(i, 7) >= brr(i, 4) Then         '判断需补货数量
            brr(i, 9) = 0
        Else
            brr(i, 9) = brr(i, 4) - brr(i, 7)
        End If
        
        brr(i, 10) = Split(brr(i, 10), ",", 2)(1) '以item number为key,将location与qty连接
    
    Next i
    Sheet1.Range("a1").CurrentRegion = brr
    Columns("a:i").EntireColumn.AutoFit
        
    Set d1 = Nothing
    Set d2 = Nothing
    Set d3 = Nothing
    Set d4 = Nothing
    
    ActiveWorkbook.Save
Application.ScreenUpdating = True
End Sub


' make transfer list for forklift guy
Sub make_list_in_sheet3()

  
    Application.ScreenUpdating = False
    Dim i&, cnn As Object, rs As Object, sql$
    Sheet3.Activate
    Sheet3.Cells.Clear
    
    Set cnn = CreateObject("adodb.connection")
    Set rs = CreateObject("ADODB.Recordset")
    
    cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=""Excel 12.0;HDR=YES"";Data Source=" & ThisWorkbook.FullName
    
    sql = "select [BDITM#],[需补货数量],[Locations]" & _
            "from [DATA$] " & _
            "where [需补货数量]<>0 "
            

    
    Set rs = cnn.Execute(sql)
    For i = 0 To rs.Fields.Count - 1
        Cells(1, i + 1) = rs.Fields(i).Name
    Next i
    
    Sheet3.Range("a2").CopyFromRecordset cnn.Execute(sql)
    
    cnn.Close
    
    Set cnn = Nothing
    Set rs = Nothing
    Sheet3.Columns("a:c").AutoFit
    Sheet3.Range("a1:d1").AutoFilter
    
    Application.ScreenUpdating = True
    
End Sub


' Excute all subs panel

Sub ReplenishFromB6toB5()
    
    t = Timer
    'ActiveWorkbook.SaveAs ActiveWorkbook.Path & "\Ashton RP Open Orders Fulfillment-" & Format(Now(), "yyyymmdd.hhmm") & ".xlsm"
    Application.ScreenUpdating = False
   
    
    Call load_TripDetails
    Call Pull_AS_STO
    Call add_LOC_QTY
    Call DATA_Columns_Ghij
    Call make_list_in_sheet3

    
    Sheet3.Range("E1").Value = "Data collected at :" & Format(Now(), "hh:mm,mm-dd-yyyy")
    Sheet3.Range("e1").Font.Color = -16776961
    Application.ScreenUpdating = True
    ThisWorkbook.Save
    MsgBox "Updated Successful~    " & Format(Timer - t, "0.00" & "s")
    
    
End Sub


' backup 
Private Sub Workbook_BeforeClose(Cancel As Boolean)

Application.DisplayAlerts = False
ans = MsgBox("Do you want to backup the file ?", vbYesNo)
If ans = vbYes Then
    ThisWorkbook.SaveAs ("D:\Document\01-Wanvog\03-Report\13-RP Orders\Fulfillment\ASRPOrderFulfillmentBackUp\AS RP ORDER FULFILLMENT  - " & Format(Now(), "yyyymmdd.hhmm") & ".xlsb")
    Else:
    Exit Sub
End If
Application.DisplayAlerts = True
End Sub








