'范例18 引用工作表的方法
'18-1 使用工作表名称
Sub shtname()
    Worksheets("Sheet10").Range("a1") = "WH"

End Sub



'18-2 使用工作表索引号
Sub ShtIndex()

    Worksheets(Worksheets.Count).Select
    
End Sub


'18-3 使用工作表代码名称
Sub ShtCodeName()
    Sheet3.Select
End Sub

'范例19 选择工作表的方法
Sub ShtSelect()
    MsgBox "下面将选择" & Sheet10.Name & "工作表"
    Sheet10.Select
    MsgBox "下面将激活" & Sheet10.Name & "工作表"
    Sheet10.Activate
End Sub

'选中所有工作表

Sub SelectSht()
    Dim sht As Worksheet
    For Each sht In Worksheets
        sht.Select False
    Next
End Sub


'20-1 使用For..Next 语句遍历工作表
Sub TraversalShtOne()
    Dim i%, Str$
    For i = 1 To Worksheets.Count
        Str = Str & Worksheets(i).Name & vbCrLf
    Next
    MsgBox " there are below worksheets: " & vbCrLf & Str
End Sub


'20-2 使用For Each...Next 语句遍历工作表

Sub TraversalShtTwo()
    Dim sht As Worksheet
    Dim Str As String
    For Each sht In Worksheets
        Str = Str & sht.Name & vbCrLf
    Next
    MsgBox " there are below worksheets: " & vbCrLf & Str

End Sub


'范例21 工作表的添加与删除

Sub ShtAddOne()

    Worksheets.Add.Name = "Data"

End Sub


Sub ShtAddTwo()
    Dim i%
    Dim sht As Worksheet
    With Worksheets
        For i = 1 To 6
            Set sht = .Add(after:=Worksheets(.Count))
            sht.Name = i
        Next
    End With
    Set sht = Nothing
End Sub


Sub ShtDel()

    Dim sht As Worksheet
    Application.DisplayAlerts = False
    For Each sht In Worksheets
        If sht.Name Like ["*班*"] Then
            sht.Delete
        End If
    Next
    Application.DisplayAlerts = True

End Sub


'先判断工作表中是否存在，再删除
Sub ShtAddThree()
    
    Dim sht As Worksheet
    For Each sht In Worksheets
        If sht.Name = "数据" Then
            If MsgBox("工作簿中已有""数据""工作表, 是否删除后添加?", 36) = 6 Then
                Application.DisplayAlerts = False
                sht.Delete
                Application.DisplayAlerts = True
            Else
                Exit Sub
            End If
        End If
    Next
    Worksheets.Add.Name = "数据"
    Set sht = Nothing
    
End Sub


Sub ShtAddFour()
    Dim arr As Variant
    Dim i As Integer
    Dim sht As Worksheet
    On Error Resume Next
    arr = Array(2, 3, 4, 5, 6, 7)
    With Worksheets
        For i = 0 To UBound(arr)
            Set sht = .Add(after:=Worksheets(.Count))
            sht.Name = arr(i)
        Next
    End With
    Application.DisplayAlerts = False
    For Each sht In Worksheets
        If sht.Name Like "Sheet*" Then sht.Delete
    Next
    Application.DisplayAlerts = True
    Set sht = Nothing
    
    
End Sub


'范例22 禁止删除指定工作表 ?????????????????  不起作用???????????????????????????
'Private Sub Workbook_activate()
'    Application.CommandBars.FindControl(ID:=847).OnAction = "MyDelSht"
'End Sub
'
'Sub MyDelSht()
'
'    If ActiveSheet.CodeName = "Sheet8" Then
'        MsgBox ActiveSheet.Name & "工作表禁止删除！", 48
'    Else
'        ActiveSheet.Delete
'    End If
'
'End Sub
'
'
'Private Sub workbook_deactivate()
'    Application.CommandBars.FindControl(ID:=847).OnAction = ""
'End Sub


'范例24 判断是否存在指定工作表

Sub ShtExists()
    
    Dim sht As Worksheet
    On Error GoTo line
    Set sht = Worksheets("abc")
    MsgBox "工作簿中已有""abc""工作表!"
    Exit Sub
line:
    MsgBox "工作薄中没有""abc""工作表!"

End Sub


''范例25 工作表的深度隐藏
'Public sht As Worksheet
'Private Sub workbook_beforeclose(Cancel As Boolean)
'    Sheet1.Visible = True
'    For Each sht In ThisWorkbook.Sheets
'        If sht.CodeName <> "Sheet1" Then
'            sht.Visible = xlSheetVeryHidden
'        End If
'    Next
'    ThisWorkbook.Save
'End Sub
'
'Private Sub workbook_open()
'    For Each sht In ThisWorkbook.Sheets
'        If sht.CodeName <> "Sheet1" Then
'            sht.Visible = xlSheetVisible
'        End If
'    Next
'    Sheet1.Visible = xlSheetVeryHidden
'
'End Sub


'范例26 工作表的保护与取消保护
Sub ShProtect()
    With Sheet13
        .Unprotect Password:="123"
        .Cells(1, 1) = .Cells(1, 1) + 100
        .Protect Password:="123"
    End With
End Sub


'VBA 解除工作表的保护

Sub RemoveShProtect()

    Dim i1%, i2%, i3%, i4%, i5%, i6%, i7%, i8%, i9%, i10%, i11%, i12%
    Dim t$
    On Error Resume Next
    If ActiveSheet.ProtectContents = False Then
        MsgBox "no pw for this workbook!"
        Exit Sub
    End If
    
    t = Timer
    For i1 = 65 To 66:  For i2 = 65 To 66: For i3 = 65 To 66
    For i4 = 65 To 66:  For i5 = 65 To 66: For i6 = 65 To 66
    For i7 = 65 To 66:  For i8 = 65 To 66: For i9 = 65 To 66
    For i10 = 65 To 66:  For i11 = 65 To 66: For i12 = 32 To 126
        ActiveSheet.Unprotect Chr(i1) & Chr(i2) & Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(i7) & Chr(i8) & Chr(i9) & Chr(i10) & Chr(i11) & Chr(i12)
        If ActiveSheet.ProtectContents = False Then
            MsgBox "解除工作表保护！ 用时" & Format(Timer - t, "0.00") & "s"
            Exit Sub
        End If
    Next: Next: Next: Next: Next: Next:
    Next: Next: Next: Next: Next: Next:
End Sub
    
    
'范例27 自动建立工作表目录 ， 选中sheet13(3)时，自动将所有sheet名汇总过来
'Private Sub worksheet_activate()
'   Dim Sht As Worksheet
'   Dim a As Integer
'   Dim r As Integer
'   r = Cells(Rows.Count, 1).End(3).Row
'   a = 2
'   If r > 1 Then Range("a2:a" & r).ClearContents
'   For Each Sht In Worksheets
'        If Sht.CodeName <> "Sheet1" Then
'            Cells(a, 1).Value = Sht.Name
'            a = a + 1
'        End If
'   Next
'   Set Sht = Nothing
'
'End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim r As Integer
    r = Cells(Rows.Count, 1).End(3).Row
    On Error Resume Next
    If Not Application.Intersect(Target, Range("a2:a" & r)) Is Nothing Then
        Sheets(Target.Text).Select
    End If
End Sub


'范例28 循环选择工作表
Sub ShtNext()
    If ActiveSheet.Index < Worksheets.Count Then
        ActiveSheet.Next.Activate
    Else
        Worksheets(1).Activate
    End If
End Sub

Sub ShtPrevious()

    If ActiveSheet.Index > 1 Then
        ActiveSheet.Previous.Activate
    Else
        Worksheets(Worksheets.Count).Activate
    End If
End Sub


'范例29 在工作表中一次插入多行
Sub InSertRow()
    Dim i As Integer
    For i = 1 To 3
        Sheet15.Rows(5).Insert
    Next
End Sub


'范例30 删除工作表中的空行

Sub DelBlankRow()
    Dim r As Long
    Dim i As Long
    r = Sheet15.UsedRange.Rows.Count
    For i = r To 1 Step -1
        If Rows(i).Find("*", , xlValues, , , 2) Is Nothing Then
        Rows(i).Delete
        End If
    Next

End Sub

'范例31 删除工作表的重复行

Sub DeleteRow()
    Dim r As Integer, i As Integer
    With Sheet15
        r = .Cells(.Rows.Count, 1).End(3).Row
        For i = r To 1 Step -1
            If WorksheetFunction.CountIf(.Columns(4), .Cells(i, 4)) > 1 Then
                .Rows(i).Delete
            End If
        Next
    End With

End Sub

'范例32 定位删除特定内容所在的行

Sub SpecialDelete()

    Dim r As Long
    With Sheet15
        r = .Cells(.Rows.Count, 1).End(3).Row
        .Range("d2:d" & r).Replace "10019262", "10019262-22222", 2
        .Columns(4).SpecialCells(4).EntireRow.Delete
        
    End With

End Sub

'范例33 判断是否选中整行
'Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'    If Target.Rows.Count = 1 Then
'        If Target.Columns.Count = 16384 Then
'            MsgBox "您选中了整行,当前行号" & Target.Row
'        End If
'    End If
'
'End Sub
'

'Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'    If Target.Rows.Count = 1048576 Then
'        If Target.Columns.Count = 1 Then
'            MsgBox "您选中了整列,当前列号" & Target.Column
'        End If
'    End If
'
'End Sub


'范例34 限制工作表的滚动区域

'Private Sub Workbook_Open()
'
'    Sheet23.ScrollArea = "$A$1:$Q$38"
'
'End Sub



'范例35 复制自动筛选后的数据区域
Sub CopyFilter()
    Sheet2.Cells.Clear
    With Sheet18
        If .FilterMode Then
            .AutoFilter.Range.SpecialCells(12).Copy Sheet2.Cells(1, 1)
        End If
    End With
End Sub


'范例36 使用高级筛选功能获得不重复记录
Sub Filter()
    Sheet18.Range("a1").CurrentRegion.AdvancedFilter Action:=xlFilterCopy, Unique:=True, copyToRange:=Sheet2.Range("a1")
    
End Sub


'范例37 获得工作表打印页数

Sub PrintPage()
    
    Dim Page As Integer
    Page = ExecuteExcel4Macro("GET.DOCUMENT(50)")
    MsgBox "工作表打印页数共" & Page & "页!"
    
End Sub

































































































