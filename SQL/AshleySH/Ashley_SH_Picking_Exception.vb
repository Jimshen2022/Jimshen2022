' Excute all subs panel

Sub Ashley_SH_Picking_Exception()
    
    t = Timer
    ActiveWorkbook.SaveAs ActiveWorkbook.Path & "\Ashley_SH_Picking_Exception -" & Format(Now(), "yyyymmdd.hhmm") & ".xlsb"
    Application.ScreenUpdating = False
   
    
    Call Pull_SN_Auto_Move
    Call Data_Column_AF
    Call Data_Column_AC
    Call Data_Column_AD_AE
    Call PivotTable_by_Reason
    Call PivotTable_by_Employee
    Call unfilter1
    
    'Sheet9.Range("a1").Value = "Data collected at :  " & Format(Now(), "hh:mm,mm-dd-yyyy")
    'Sheet9.Range("a1").Font.Color = -16776961
    Application.ScreenUpdating = True
    ThisWorkbook.Save
    MsgBox "Updated Successful~    " & Format(Timer - t, "0.00" & "s")
    
    
End Sub




Sub Pull_SN_Auto_Move()
    't = Timer
    Application.ScreenUpdating = False
    Dim wb As Workbook
    Dim arr, brr(), i&, j&, k&, nrow&, crr()
    Sheets("SETUP").Range("A1").Value = "Data collected at:" & Format(Now(), "hh:mm,mm-dd-yyyy")
    Sheets("DATA").Select
    Cells.Delete
    Set wb = GetObject("C:\Users\jishen\Downloads\840.xlsx")
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False
'    ReDim brr(1 To UBound(arr), 1 To UBound(arr, 2))
    
'    For i = 1 To UBound(arr)
'            brr(i, 1) = arr(i, 2)
'            brr(i, 2) = arr(i, 3)
'            brr(i, 3) = arr(i, 7)
'
'        For j = 1 To UBound(arr, 2)
'            brr(i, j) = arr(i, j)
'        Next
'    Next
    
    Columns("a:ab").NumberFormat = "@"
    Sheets("DATA").Range("a1").Resize(UBound(arr), UBound(arr, 2)) = arr
    Columns("a:ab").EntireColumn.AutoFit
    ActiveWorkbook.Save
    Application.ScreenUpdating = True
    'MsgBox Format(Timer - t, "0.00" & "s")

End Sub


Sub Data_Column_AF()  ' Sheet DATA "Team"
    't = Timer
    Application.ScreenUpdating = False
    Dim wb As Workbook
    Dim arr(), brr(), i&
    Dim d As Object
    Set d = CreateObject("scripting.dictionary")
    
    Sheets("Team").Activate
    brr = Sheet5.Range("a1").CurrentRegion
        For i = 2 To UBound(brr)
            d(brr(i, 1)) = brr(i, 2)
        Next
    
    Sheets("DATA").Activate
    Range("ac1:ag1").Value = Array("Reason", "Type", "IssueType", "Team", "Remark")
    arr = Sheet4.Range("a1").CurrentRegion
        
    For i = 2 To UBound(arr) '遍历查询值
        If d.Exists(arr(i, 18)) Then '如果字典存在查询值
            arr(i, 32) = d(arr(i, 18)) '获取人名对应的条目
        Else
            arr(i, 32) = "请至Sheet-Team维护员工姓名先"
        End If
    Next
'   Range("a1:ag" & Cells(Rows.Count, "a").End(xlUp).Row) = arr
    Columns("a:ag").NumberFormat = "@"
    Range("a1").Resize(UBound(arr), UBound(arr, 2)) = arr
    Columns("a:ab").EntireColumn.AutoFit
    Set d = Nothing
    ActiveWorkbook.Save
    Application.ScreenUpdating = True
    'MsgBox Format(Timer - t, "0.00" & "s")
End Sub


Sub Data_Column_AC()  ' Sheet DATA "Reason"
    't = Timer
    Application.ScreenUpdating = False
    Dim wb As Workbook
    Dim arr(), i&
    
    Sheets("DATA").Activate
    arr = Sheet4.Range("a1").CurrentRegion
        
    For i = 2 To UBound(arr)
        If arr(i, 24) = "800" Then
            arr(i, 29) = "Cycle Count"
        ElseIf arr(i, 7) Like "[QA,FX,EN]*" Then
            arr(i, 29) = "OBQ Returned"
        ElseIf (arr(i, 7) Like "K*" And (arr(i, 9) Like "K*" Or arr(i, 9) Like "ST*" Or arr(i, 9) Like "FL*") And arr(i, 24) Like "3*") Then
            arr(i, 29) = "Picked from A location, but SN in B location"
        ElseIf (arr(i, 7) Like "K*" And (arr(i, 9) Like "K*" Or arr(i, 9) Like "ST*" Or arr(i, 9) Like "FL*") And arr(i, 24) Like "2*") Then
            arr(i, 29) = "Moved from A location, but SN in B location"
        ElseIf arr(i, 7) Like "K*" And (arr(i, 9) Like "ST*" Or arr(i, 9) Like "V*" Or arr(i, 9) Like "S*" Or arr(i, 9) Like "F*") And arr(i, 24) Like "3*" Then
            arr(i, 29) = "Didn't pick, took physical to stage and loaded"
        ElseIf arr(i, 7) Like "K*" And (arr(i, 9) Like "ST*" Or arr(i, 9) Like "V*" Or arr(i, 9) Like "S*" Or arr(i, 9) Like "F*") And arr(i, 24) Like "2*" Then
            arr(i, 29) = "Moved from A location, but SN in B location"
        ElseIf arr(i, 7) Like "K*" And (arr(i, 9) Like "ST*" Or arr(i, 9) Like "V*" Or arr(i, 9) Like "S*" Or arr(i, 9) Like "F*") And arr(i, 24) Like "1*" Then
            arr(i, 29) = "Receiving issue"
        ElseIf (arr(i, 7) Like "FL*" Or arr(i, 7) Like "ST*" Or arr(i, 7) Like "PS*" Or arr(i, 7) Like "EN*") And (arr(i, 9) Like "V*" Or arr(i, 9) Like "K*" Or arr(i, 9) Like "ST*" Or arr(i, 9) Like "FL*") Then
            arr(i, 29) = "Picked from floppy location"
        ElseIf arr(i, 7) Like "S0*" And (arr(i, 9) Like "S2*" Or arr(i, 9) Like "S5*" Or arr(i, 9) Like "S6*") Then
            arr(i, 29) = "Didn't do hotload move but loading scan"
        ElseIf arr(i, 7) Like "S*" And (arr(i, 9) Like "V*" Or arr(i, 9) Like "F*" Or arr(i, 9) Like "ST*" Or arr(i, 9) Like "K*") And arr(i, 24) Like "3*" Then
            arr(i, 29) = "Picked from stage location"
        ElseIf arr(i, 7) Like "S*" And (arr(i, 9) Like "V*" Or arr(i, 9) Like "F*" Or arr(i, 9) Like "ST*" Or arr(i, 9) Like "K*") And arr(i, 24) Like "2*" Then
            arr(i, 29) = "Moved from stage location"
        ElseIf arr(i, 7) Like "WR*" And (arr(i, 9) Like "FX*" Or arr(i, 9) Like "F*" Or arr(i, 9) Like "ST*" Or arr(i, 9) Like "K*") And arr(i, 24) Like "2*" Then
            arr(i, 29) = "Inbound Damaged Moving"
        ElseIf arr(i, 7) Like "RS*" And (arr(i, 9) Like "ST*" Or arr(i, 9) Like "FL*" Or arr(i, 9) Like "K*" Or arr(i, 9) Like "V*") And arr(i, 24) Like "2*" Then
            arr(i, 29) = "Didn't direct pickup but moved from RS318AA1"
        ElseIf arr(i, 7) Like "RP*" And arr(i, 9) Like "RP*" And arr(i, 24) Like "2*" Then
            arr(i, 29) = "RPC Cycle count stage location"
        ElseIf arr(i, 7) Like "S0*" And arr(i, 9) Like "ST*" And arr(i, 10) Like "S*" Then
            arr(i, 29) = "Missed move from big stage to small stage but loaded directly"
        Else
            arr(i, 29) = "Please check the reason"
        End If
    Next
    Columns("a:ag").NumberFormat = "@"
    Range("a1:ag" & Cells(Rows.Count, "a").End(xlUp).Row) = arr
'    Range("a1").Resize(UBound(arr), UBound(arr, 2)) = arr
    Columns("a:ab").EntireColumn.AutoFit
    'ActiveWorkbook.Save
    Application.ScreenUpdating = True
    'MsgBox Format(Timer - t, "0.00" & "s")

End Sub


Sub Data_Column_AD_AE()  ' Sheet DATA "Type","IssueType"
    't = Timer
    Application.ScreenUpdating = False
    Dim wb As Workbook
    Dim arr(), brr(), i&
    Dim d As Object, d2 As Object
    Set d = CreateObject("scripting.dictionary")
    Set d2 = CreateObject("scripting.dictionary")
    
    Sheets("Code").Activate
    brr = Sheet6.Range("a1").CurrentRegion
        For i = 2 To UBound(brr)
            d(brr(i, 1)) = brr(i, 2)   'Type
            d2(brr(i, 1)) = brr(i, 3)  ' Issue Type
        Next
    
    Sheets("Data").Activate
    arr = Sheet4.Range("a1").CurrentRegion
        
    For i = 2 To UBound(arr)
        If d.Exists(arr(i, 29)) Then '如果字典存在查询值
            arr(i, 30) = d(arr(i, 29)) '获取人名对应的条目
        Else
            arr(i, 30) = "请确认异常原因"
        End If
        If d2.Exists(arr(i, 29)) Then '如果字典存在查询值
            arr(i, 31) = d2(arr(i, 29)) '获取人名对应的条目
        Else
            arr(i, 31) = "请确认异常原因"
        End If
        
    Next
    Columns("a:ag").NumberFormat = "@"
    Range("a1:ag" & Cells(Rows.Count, "a").End(xlUp).Row) = arr
    'Range("a1").Resize(UBound(arr), UBound(arr, 2)) = arr
    Columns("a:ag").EntireColumn.AutoFit
    Set d = Nothing
    Set d2 = Nothing
    'ActiveWorkbook.Save
    Application.ScreenUpdating = True
    'MsgBox Format(Timer - t, "0.00" & "s")
End Sub

Sub PivotTable_by_Reason()

    't = Timer
    Application.ScreenUpdating = False
    Dim wksData As Worksheet
    Dim objCache As PivotCache
    Dim objTabel As PivotTable
    Dim arr()
    
    With Sheet3   '需建立的 PivotTable所在sheet
        .Cells.Clear
        Set wksData = Sheet4
        arr = wksData.Range("a1").CurrentRegion
        Set objCache = ThisWorkbook.PivotCaches.Create(xlDatabase, wksData.Range("a1").CurrentRegion.Address(external:=True))
        Set objTabel = objCache.CreatePivotTable(Sheet3.Range("a3"))     '建立pivottable的位置
        With objTabel
             .AddFields RowFields:=Array(arr(1, 30), arr(1, 29)), _
                    ColumnFields:=Array(arr(1, 20)), _
                    PageFields:=Array(arr(1, 31))
             .AddDataField .PivotFields(arr(1, 5)), , xlCount '
             .RowAxisLayout xlOutlineRow
            .MergeLabels = True               '如下3个为合并单元格并采用传统显示模式
            .InGridDropZones = True
            .RowAxisLayout xlTabularRow
'           .PivotFields("IssueType").PivotItems("None picking exception").Visible = False   '左上角筛选issue type --- 间接选中,排除其他的方式
            .PivotFields("IssueType").CurrentPage = "picking exception"                     '左上角筛选issue type --- 直接选中的方式
            .PivotFields("Type").Subtotals(1) = False     '取消Type的subtotal
        End With
        .Columns("a:a").ColumnWidth = 20.14
        
    End With
    Sheet3.Range("e1").Value = "DataCollectedAt:  " & Format(Now, "HH:MM:SSam/pm,mmm.dd.yyyy")
    Sheet3.Range("e1").Font.Color = -16776961
    
    'ActiveWorkbook.Save
    
    Application.ScreenUpdating = True
    'MsgBox Format(Timer - t, "0.00" & "s")
    
End Sub


Sub PivotTable_by_Employee()

    't = Timer
    Application.ScreenUpdating = False
    Dim wksData As Worksheet
    Dim objCache As PivotCache
    Dim objTabel As PivotTable
    Dim arr()
    
    With Sheet2   '需建立的 PivotTable所在sheet
        .Cells.Clear
        Set wksData = Sheet4
        arr = wksData.Range("a1").CurrentRegion
        Set objCache = ThisWorkbook.PivotCaches.Create(xlDatabase, wksData.Range("a1").CurrentRegion.Address(external:=True))
        Set objTabel = objCache.CreatePivotTable(Sheet2.Range("a3"))     '建立pivottable的位置
        With objTabel
             .AddFields RowFields:=Array(arr(1, 18), arr(1, 32)), _
                    ColumnFields:=Array(arr(1, 20)), _
                    PageFields:=Array(arr(1, 31))
             .AddDataField .PivotFields(arr(1, 5)), , xlCount '
             .RowAxisLayout xlOutlineRow
            .MergeLabels = True               '如下3个为合并单元格并采用传统显示模式
            .InGridDropZones = True
            .RowAxisLayout xlTabularRow
'           .PivotFields("IssueType").PivotItems("None picking exception").Visible = False   '左上角筛选issue type --- 间接选中,排除其他的方式
            .PivotFields("IssueType").CurrentPage = "picking exception"                     '左上角筛选issue type --- 直接选中的方式
            .PivotFields("Empname").Subtotals(1) = False     '取消Empname的subtotal
             
        End With
        .Columns("a:a").ColumnWidth = 20.14
        
    End With
    Sheet2.Range("e1").Value = "DataCollectedAt:  " & Format(Now, "HH:MM:SSam/pm,mmm.dd.yyyy")
    Sheet2.Range("e1").Font.Color = -16776961
    
    'ActiveWorkbook.Save
    
    Application.ScreenUpdating = True
    'MsgBox Format(Timer - t, "0.00" & "s")
    
End Sub

Sub unfilter1()
    '取消筛选
    Application.ScreenUpdating = False
    Dim sht As Worksheet
    Sheet4.Select
    

       With Range("A1:AG1")
            .Interior.ColorIndex = 49
            .Font.ColorIndex = 2
            Range("B2").Select
            ActiveWindow.FreezePanes = True
            .Columns("A:AG").EntireColumn.AutoFit
        End With
        
    Set sht = Sheet4
        ' 如果当前工作表为筛选模式，则取消
        If sht.AutoFilterMode = True Then sht.AutoFilterMode = 0
        ' 如果当前工作表没有筛选，则加上筛选
        If sht.AutoFilterMode = False Then sht.Range("a1").AutoFilter Field:=1
     Application.ScreenUpdating = True
        
End Sub








