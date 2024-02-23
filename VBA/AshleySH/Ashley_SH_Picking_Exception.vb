' Excute all subs panel

Sub AS_HJ_vs_Mapics_Variance_Report()
    
    t = Timer
    Application.ScreenUpdating = False
    ActiveWorkbook.SaveAs ActiveWorkbook.Path & "\AS HJvsMapics Variance Report-" & Format(Now(), "yyyymmdd.hhmm") & ".xlsb"
'    Sheet9.Cells.Clear
'    Sheet21.Range("a1:ac" & Rows.Count).Copy
'    Sheet9.Range("a1").PasteSpecial Paste:=xlPasteValues
'    Application.CutCopyMode = False
'
    Call load_WH22_OnHand
    Call load_IA_TRX
    Call Pull_HJ_vs_Mapics
    Call Pull_Search_STO_and_SNA_Balance
    Call Pull_SN_HOLD
    Call DATA_ColumnC_Reason
    Call unfilter1
    Call DATA_Columns_NOPQRS
    
    'Sheet3.Range("E1").Value = "Data collected at :" & Format(Now(), "hh:mm,mm-dd-yyyy")
    Application.ScreenUpdating = True
    ThisWorkbook.Save
    Sheet21.Select
    MsgBox "Updated Successful~    " & Format(Timer - t, "0.00" & "s")
        
End Sub



'Pull WH22
Sub load_WH22_OnHand()
    
    Application.ScreenUpdating = False
    Dim i As Long
    Dim adors As New Recordset
    Sheet12.Activate
    Sheet12.Cells.Clear
    
    
    Set Db = New Connection
    Db.CursorLocation = adUseClient
    If Db.State = 1 Then Db.Close
    
    U = Sheet2.Range("a1").Value
    P = Sheet2.Range("a2").Value
    
    Db.Open "Provider =IBMDASQL.DataSource.1" & _
            ";Catalog Library List=JIMTDTA" & _
            ";Persist Security Info=True" & _
            ";Force Translate=0" & _
            ";Data Source = WVFHA" & _
            ";User ID =" & U & "" & _
            ";Password =" & P
    
    Set adors = New Recordset
    If adors.State = 1 Then adors.Close
    
     cmdtxt = "Select t1.ITNBR, t1.HOUSE, t1.ITCLS, t1.MOHTQ, t1.WHSLC, t1.QTSYR, t2.ITDSC, t4.LLOCN, t4.LQNTY " & _
              "from AMFLIBQ.ITEMBL t1, AMFLIBQ.ITMRVA t2, AMFLIBQ.WHSMST t3, AMFLIBQ.SLQNTY t4 " & _
              "where t1.ITNBR = t4.ITNBR AND t4.ITNBR = t2.ITNBR AND t1.ITCLS= t2.ITCLS AND t1.HOUSE=t4.HOUSE AND t4.HOUSE = t3.WHID AND t3.STID = t2.STID AND ((t1.MOHTQ<>0) AND (t1.HOUSE='22')) " & _
              "order by t1.ITNBR "

'SELECT t1.ITNBR, t1.HOUSE, t1.ITCLS, t1.MOHTQ, t1.WHSLC, t1.QTSYR, t2.ITDSC, t4.LLOCN, t4.LQNTY
'FROM AMFLIBQ.ITEMBL t1, AMFLIBQ.ITMRVA t2, AMFLIBQ.WHSMST t3, AMFLIBQ.SLQNTY t4
'WHERE t1.ITNBR = t4.ITNBR AND t4.ITNBR = t2.ITNBR AND t1.ITCLS= t2.ITCLS AND t1.HOUSE=t4.HOUSE AND t4.HOUSE = t3.WHID AND t3.STID = t2.STID AND ((t1.MOHTQ<>0) AND (t1.HOUSE='22'))
'ORDER BY t1.ITNBR
'
    adors.Open cmdtxt, Db, 3, 3
    For i = 0 To adors.Fields.Count - 1
        Sheet12.Cells(1, i + 1) = adors.Fields(i).Name
    Next i
    
    Sheet12.Columns("A:C").NumberFormat = "@"
    Sheet12.Columns("e:h").NumberFormat = "@"
    Sheet12.Range("a2").CopyFromRecordset adors
    adors.Close
    Set adors = Nothing
    

    
    Application.ScreenUpdating = True
End Sub

'Pull IA transactions
Sub load_IA_TRX()
    
    Application.ScreenUpdating = False
    Dim i As Long
    Dim adors As New Recordset
    Sheet7.Activate
    Sheet7.Cells.Clear
    
    
    Set Db = New Connection
    Db.CursorLocation = adUseClient
    If Db.State = 1 Then Db.Close
    
    U = Sheet2.Range("a1").Value
    P = Sheet2.Range("a2").Value
    
    Db.Open "Provider =IBMDASQL.DataSource.1" & _
            ";Catalog Library List=JIMTDTA" & _
            ";Persist Security Info=True" & _
            ";Force Translate=0" & _
            ";Data Source = WVFHA" & _
            ";User ID =" & U & "" & _
            ";Password =" & P
    
    Set adors = New Recordset
    If adors.State = 1 Then adors.Close
    
     cmdtxt = "SELECT t1.TCODE, t1.ORDNO, t1.ITNBR, t2.ITCLS, t1.HOUSE, t1.UPDDT, t1.UPDTM, t1.TRQTY, t1.TRNDT, t1.LBHNO, t1.REFNO, t1.REASN, t1.USRSQ " & _
              "FROM AMFLIBQ.IMHIST t1, AMFLIBQ.ITMRVA t2, AMFLIBQ.WHSMST t3 " & _
              "WHERE t2.ITNBR = t1.ITNBR AND t2.STID = t3.STID AND t1.HOUSE = t3.WHID AND ((t1.HOUSE='232') AND " & _
              "(t1.UPDDT BETWEEN  '1210101' AND int('1'||substr(trim(char(CURRENT DATE)),3,2)||substr(trim(char(CURRENT DATE)),6,2)||substr(trim(char(CURRENT DATE)),9,2))) " & _
              "AND (t1.TRQTY<>0) AND t1.TCODE IN ('IA','SS','RC')) "

    adors.Open cmdtxt, Db, 3, 3
    For i = 0 To adors.Fields.Count - 1
        Sheet7.Cells(1, i + 1) = adors.Fields(i).Name
    Next i
    
    Sheet7.Columns("A:G").NumberFormat = "@"
    Sheet7.Columns("I:M").NumberFormat = "@"
    Sheet7.Range("a2").CopyFromRecordset adors
    Columns("a:m").EntireColumn.AutoFit
    adors.Close
    Set adors = Nothing
    

    
    Application.ScreenUpdating = True
End Sub


'Pull_HJ VS Mapics_Variance_REPORT

Sub Pull_HJ_vs_Mapics()
    Dim wb As Workbook
    Dim arr, brr(), i&, j&, k&, nrow&, crr()
        
    't = Timer
    Application.ScreenUpdating = False
    
    Sheet21.Select
    Cells.Clear
    
    
    Set wb = GetObject("C:\Users\jishen\Downloads\Mapics_vs.xlsx")
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False
    ReDim brr(2 To UBound(arr), 1 To UBound(arr, 2))

    For i = 2 To UBound(arr)
'            brr(i, 1) = arr(i, 2)
'            brr(i, 2) = arr(i, 3)
'            brr(i, 3) = arr(i, 7)
        For j = 1 To UBound(arr, 2)
            brr(i, j) = arr(i, j)
        Next
    Next
    
    Columns("a:b").NumberFormat = "@"
    'Columns("c:d").NumberFormat = "@"
    Sheet21.Range("a1").Resize(UBound(arr) - 1, UBound(arr, 2)) = brr
    Range("z1").Value = arr(1, 1)
    Range("c:d").EntireColumn.Insert
    Range("c1:d1").Value = Array("Reason", "ABS")
    
    crr = Range("a1").CurrentRegion
    For k = 2 To UBound(crr)
        crr(k, 4) = Abs(crr(k, 5))
    Next
    Range("a1").Resize(UBound(crr), UBound(crr, 2)).Value = crr
    
    Range("d2:m" & UBound(crr)).Value = Range("d2:m" & UBound(crr)).Value
    Columns("a:ab").EntireColumn.AutoFit
    
    Application.ScreenUpdating = True
    'MsgBox Format(Timer - t, "0.00" & "s")
    
End Sub

'Pull STO and SNA Balance Report


Sub Pull_Search_STO_and_SNA_Balance()
    Dim wb As Workbook
    Dim arr, brr(), i&, j&, k&, nrow&, crr()
        
    't = Timer
    Application.ScreenUpdating = False
    
    Sheet8.Select
    Cells.Clear
    
    
    Set wb = GetObject("C:\Users\jishen\Downloads\Search_STO_and_SNA_Balance.xlsx")
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False
    ReDim brr(1 To UBound(arr), 1 To UBound(arr, 2))

    For i = 1 To UBound(arr)
'            brr(i, 1) = arr(i, 2)
'            brr(i, 2) = arr(i, 3)
'            brr(i, 3) = arr(i, 7)
        For j = 1 To UBound(arr, 2)
            brr(i, j) = arr(i, j)
        Next
    Next
    
    Columns("a:c").NumberFormat = "@"
    'Columns("c:d").NumberFormat = "@"
    Sheet8.Range("a1").Resize(UBound(arr), UBound(arr, 2)) = brr
    
    Range("a1").Resize(UBound(arr), UBound(arr, 2)).Value = brr
    
    Range("d2:g" & UBound(arr)).Value = Range("d2:g" & UBound(arr)).Value
    Columns("a:g").EntireColumn.AutoFit
    
    Application.ScreenUpdating = True
    'MsgBox Format(Timer - t, "0.00" & "s")
    
End Sub


'Pull SN HOLD

Sub Pull_SN_HOLD()

    Dim wb As Workbook
    Dim arr, brr(), i&, j&, k&, m&, nrow&, crr()
    
    't = Timer
    Application.ScreenUpdating = False
    
    Sheet14.Activate
    Cells.Clear
    
    
    Set wb = GetObject("C:\Users\jishen\Downloads\AS_HOLD.xlsx") '打开工作簿
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False
    ReDim brr(1 To UBound(arr), 1 To UBound(arr, 2) + 1)
    For i = 1 To UBound(arr)
        
        If arr(i, 9) <> "QA001VD1" And arr(i, 5) <> "Orphaned" Then
            m = m + 1
            For j = 1 To 16
                brr(m, j) = arr(i, j)
                brr(m, 17) = 1
                
            Next
            
        End If
    Next
    Columns("a:p").NumberFormat = "@"
    brr(1, 17) = "Qty"

    Sheet14.Range("a1").Resize(UBound(arr), UBound(arr, 2) + 1) = brr
    Columns("a:p").EntireColumn.AutoFit
    
    Application.ScreenUpdating = True
    'MsgBox Format(Timer - t, "0.00" & "s")
    
End Sub


'Pull SN ORPHANED

Sub Pull_SN_ORPHANED()

    Dim wb As Workbook
    Dim arr, brr(), i&, j&, k&, m&, nrow&, crr()
    
    't = Timer
    Application.ScreenUpdating = False
    
    Sheet4.Activate
    Cells.Clear
        
    Set wb = GetObject("C:\Users\jishen\Downloads\AS_ORPHANED.xlsx") '打开工作簿
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False
       
    ReDim brr(1 To UBound(arr), 1 To 6)
        For i = 1 To UBound(arr)
            
            brr(i, 1) = arr(i, 2)
            brr(i, 2) = arr(i, 3)
            brr(i, 3) = arr(i, 5)
            brr(i, 4) = arr(i, 9)
            brr(i, 5) = arr(i, 11)
            brr(i, 6) = arr(i, 8)
        Next
        
        With Sheet4
         .Columns("a:f").NumberFormat = "@"
         .Range("a1").Resize(UBound(arr), 6) = brr
         .Columns("a:f").EntireColumn.AutoFit
        End With
    Application.ScreenUpdating = True
    'MsgBox Format(Timer - t, "0.00" & "s")
    
End Sub

Sub DATA_ColumnC_Reason()

    Application.ScreenUpdating = False
    Dim d As Object
    Dim i&, arr(), brr()
    Set d1 = CreateObject("scripting.dictionary")
    arr = Sheet9.Range("a1").CurrentRegion 'Data Source as dictionary
    For i = 2 To UBound(arr)
        d1(arr(i, 2)) = arr(i, 3)
    Next i
    brr = Sheet21.Range("a1").CurrentRegion
    For i = 2 To UBound(brr)
        If d1.exists(brr(i, 2)) Then
            brr(i, 3) = d1(brr(i, 2))
        Else
            brr(i, 3) = "Null"
        End If
    Next i
    Sheet21.Range("a1").CurrentRegion = brr
    'Columns("a:i").EntireColumn.AutoFit
    Set d1 = Nothing
    ActiveWorkbook.Save
Application.ScreenUpdating = True
End Sub


Sub unfilter1()
    '取消筛选
    Application.ScreenUpdating = False
    Dim sht As Worksheet
    Sheet21.Select
    

       With Range("A1:Y1")
            .Range("N1:Y1").Value = Array("LQ01", "QA01", "RD01", "FA00", "RP01", "GAP", "1459", "SNA", "RC", "SS", "HOLD", "HOLD+DIFF")
            .Interior.ColorIndex = 49
            .Font.ColorIndex = 2
            Range("C2").Select
            ActiveWindow.FreezePanes = True
            .Columns("D:Y").EntireColumn.AutoFit
        End With
        
    Set sht = Sheet21
        ' 如果当前工作表为筛选模式，则取消
        If sht.AutoFilterMode = True Then sht.AutoFilterMode = 0
        ' 如果当前工作表没有筛选，则加上筛选
        If sht.AutoFilterMode = False Then sht.Range("a1").AutoFilter Field:=1
     Application.ScreenUpdating = True
        
End Sub

Sub DATA_Columns_NOPQRS()

    Application.ScreenUpdating = False
    Dim d As Object
    Dim i&, j&, k&, m&, n&, arr(), brr(), crr(), drr(), err(), frr()
    Set d1 = CreateObject("scripting.dictionary")
    Set d2 = CreateObject("scripting.dictionary")
    Set d3 = CreateObject("scripting.dictionary")
    Set d4 = CreateObject("scripting.dictionary")
    Set d5 = CreateObject("scripting.dictionary")
    Set d6 = CreateObject("scripting.dictionary")
    Set d7 = CreateObject("scripting.dictionary")
    Set d8 = CreateObject("scripting.dictionary")
    Set d9 = CreateObject("scripting.dictionary")
    Set d10 = CreateObject("scripting.dictionary")
    Set d11 = CreateObject("scripting.dictionary")
    arr = Sheet12.Range("a1").CurrentRegion '需要汇总的来源表WH22
    crr = Sheet6.Range("a1").CurrentRegion '需要汇总的来源表1495PCS
    drr = Sheet8.Range("a1").CurrentRegion '需要汇总的来源表SNA
    err = Sheet7.Range("a1").CurrentRegion        'source table IA
    frr = Sheet14.Range("a1").CurrentRegion       'source table HOLD
    For i = 2 To UBound(arr)
        If arr(i, 8) = "LQ01" Then
            d1(arr(i, 1)) = d1(arr(i, 1)) + arr(i, 9) '以item number为key,将Qty累加到字典1
        ElseIf arr(i, 8) = "QA01" Then
            d2(arr(i, 1)) = d2(arr(i, 1)) + arr(i, 9) '以item number为key,将Qty累加到字典2
        ElseIf arr(i, 8) = "RD01" Then
            d3(arr(i, 1)) = d3(arr(i, 1)) + arr(i, 9) '以item number为key,将Qty累加到字典2
        ElseIf arr(i, 8) = "FA00" Then
            d4(arr(i, 1)) = d4(arr(i, 1)) + arr(i, 9) '以item number为key,将Qty累加到字典2
        ElseIf arr(i, 8) = "RP01" Then
            d5(arr(i, 1)) = d5(arr(i, 1)) + arr(i, 9) '以item number为key,将Qty累加到字典2
        'd(arr(i, 2)) = d(arr(i, 2)) & " " & arr(i, 7) & "*" & arr(i, 3) & "pcs" '以item number为key,将location与qty连接
        Else
            d6(arr(i, 1)) = d6(arr(i, 1)) + arr(i, 9)   '以item number为key,将Qty累加到字典2
        End If
    Next i
    
    For j = 2 To UBound(crr)
        d7(crr(j, 2)) = d7(crr(j, 2)) + crr(i, 13)    ' load 1495pcs into d7
    Next j
    
    For k = 2 To UBound(drr)
        d8(drr(k, 2)) = d8(drr(k, 2)) + crr(i, 7)    ' load SNA into d8
    Next k
    
    For m = 2 To UBound(err)
        If err(m, 1) = "RC" Then d9(err(m, 3)) = d9(err(m, 3)) + err(m, 8)    ' load RC into dictionary d9
        If err(m, 1) = "SS" Then d10(err(m, 3)) = d10(err(m, 3)) + err(m, 8)  ' load SS into dictionary d10
    
    Next m
    
    For n = 2 To UBound(frr)
         d11(frr(n, 3)) = d11(frr(n, 3)) + frr(n, 17)    ' load hold sn into d11
    Next n
    
    
    brr = Sheet21.Range("a1").CurrentRegion
    For i = 2 To UBound(brr)
        If d1.exists(brr(i, 2)) Then
            brr(i, 14) = d1(brr(i, 2))
        Else
                brr(i, 14) = 0          '
        End If
        If d2.exists(brr(i, 2)) Then
            brr(i, 15) = d2(brr(i, 2))
        Else
                brr(i, 15) = 0
        End If
        
        If d3.exists(brr(i, 2)) Then
            brr(i, 16) = d3(brr(i, 2))
        Else
            brr(i, 16) = 0
        End If
        
        If d4.exists(brr(i, 2)) Then
            brr(i, 17) = d4(brr(i, 2))
        Else
                brr(i, 17) = 0
        End If
        
        If d5.exists(brr(i, 2)) Then
            brr(i, 18) = d5(brr(i, 2))
        Else
                brr(i, 18) = 0
        End If
        brr(i, 19) = brr(i, 14) + brr(i, 15) + brr(i, 16) + brr(i, 17) + brr(i, 18) + brr(i, 5) 'GAP
        
        If d7.exists(brr(i, 2)) Then
            brr(i, 20) = d7(brr(i, 2))
        Else
            brr(i, 20) = 0
        End If

        If d8.exists(brr(i, 2)) Then
            brr(i, 21) = d8(brr(i, 2))
        Else
            brr(i, 21) = 0
        End If

        If d9.exists(brr(i, 2)) Then
            brr(i, 22) = d9(brr(i, 2))
        Else
            brr(i, 22) = 0
        End If
        If d10.exists(brr(i, 2)) Then
            brr(i, 23) = d10(brr(i, 2))
        Else
            brr(i, 23) = 0
        End If
        
        If d11.exists(brr(i, 2)) Then
            brr(i, 24) = d11(brr(i, 2))
        Else
            brr(i, 24) = 0
        End If
        
        brr(i, 25) = brr(i, 24) + brr(i, 5)

    Next i
    
    Sheet21.Activate
    Sheet21.Range("a1").CurrentRegion = brr

    Columns("d:y").EntireColumn.AutoFit
    With Range("n1:y1")
        .Interior.ColorIndex = 43
        .Font.ColorIndex = 1
    End With
    Range("s1").Interior.ColorIndex = 27
    Range("ac1").Font.ColorIndex = 3
    Range("f2:f" & UBound(brr)).Font.ColorIndex = 3
    Range("i2:i" & UBound(brr)).Font.ColorIndex = 3
    Range("s2:s" & UBound(brr)).Font.ColorIndex = 3
        
    Set d1 = Nothing
    Set d2 = Nothing
    Set d3 = Nothing
    Set d4 = Nothing
    Set d5 = Nothing
    Set d6 = Nothing
    Set d7 = Nothing
    Set d8 = Nothing
    Set d9 = Nothing
    Set d10 = Nothing
    Set d11 = Nothing
            
    Application.ScreenUpdating = True
End Sub







