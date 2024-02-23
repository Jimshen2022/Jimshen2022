'第1章  Range(单元格)对象
'范例1 单元格的引用方法
'1-1 使用Range属性引用单元格区域

Sub MyRng()
    Range("A1:B4,D5:E8").Select
    Range("a1").Formula = "=rand()"
    Range("a1:b4 b2:b6").Value = 10
    Range("a1", "b4").Font.Italic = True
    
End Sub

'1-2 使用Cells属性引用单元格区域
Sub MyCell()
    Dim i As Byte
    For i = 1 To 10
        Sheets("jim").Cells(i, 1).Value = i
    Next
    Range(Cells(1, 1), Cells(10, 2)).Interior.Color = 5610
End Sub


'1-3 使用快捷记号实现快速输入
Sub FastMark()
    [a1] = "Excel 2007"
    [rng] = 1000
End Sub

'1-4 使用Offset属性返回单元格区域
Sub RngOffset()

    Sheets("jim").Range("a11:b12").Offset(2, 2).Select

End Sub

'1-5 使用Resize属性返回调整后的单元格区域
Sub RngResize()

    Sheets("jim").Range("a1").Resize(4, 4).Select

End Sub


'范例2 选定单元格区域的方法
'2-1 使用Select方法选定单元格区域
Sub RngSelect()

    Sheets("jim").Activate
    Sheets("jim").Range("a1:b10").Select
    
End Sub

'2-2 使用Activate方法选定单元格区域
Sub RngActivate()
    
    Sheets("jim").Activate
    Sheets("jim").Range("a1:x10").Activate

End Sub


'2-3 使用Go方法选定单元格区域
Sub RngGoto()

    Application.Goto reference:=Sheets("jim").Range("a1000:d1000"), Scroll:=True

End Sub


'范例3 获得指定行的最后一个非空单元格

Sub LastCell()

    Dim rng As Range, rng2 As Range
    Set rng = Cells(Rows.Count, 1).End(xlUp)
    Set rng2 = Cells(1, Columns.Count).End(1)
    MsgBox "a列的最后一个非空单元格是" & rng.Address(0, 0) & ", 行号" & rng.Row & ",数值" & rng.Value
    MsgBox "第1行的最后一个非空单元格是" & rng2.Address(0, 0) & ", 行号" & rng2.Row & ",数值" & rng2.Value
    Set rng = Nothing
    
End Sub

'范例4 使用SpecialCells方法定位单元格   '查找已使用单元格区域中含有公式的单元格
Sub SpecialAddress()

    Dim rng As Range
    Set rng = Sheet27.UsedRange.SpecialCells(xlCellTypeFormulas)
    rng.Select
    MsgBox "工作表中有公式的单元格为: " & rng.Address
    Set rng = Nothing

End Sub


'范例5 查找特定内容的单元格
'5-1 Find
Sub FindCell()

    Dim StrFind As String
    Dim rng As Range
    StrFind = InputBox("请输入要查找的值: ")
    If Len(Trim(StrFind)) > 0 Then
        With Sheet1.Range("a:b")
            Set rng = .Find(what:=StrFind, _
            after:=.Cells(.Cells.Count), _
            LookIn:=xlValues, _
            lookat:=xlWhole, _
            Searchorder:=xlByRows, _
            SearchDirection:=xlNext, _
            MatchCase:=False)
            If Not rng Is Nothing Then
                Application.Goto rng, True
            Else
                MsgBox "nothing be found!"
            End If
        End With
    End If
    Set rng = Nothing
    
End Sub



Sub FindNextCell()

    Dim StrFind As String
    Dim rng As Range
    Dim FindAddress As String
    
    StrFind = InputBox("请输入要查找的值: ")
    If Len(Trim(StrFind)) > 0 Then
        With Sheet1.Range("a:a")
            .Interior.ColorIndex = 0
            Set rng = .Find(what:=StrFind, _
            after:=.Cells(.Cells.Count), _
            LookIn:=xlValues, _
            lookat:=xlWhole, _
            Searchorder:=xlByRows, _
            SearchDirection:=xlNext, _
            MatchCase:=False)
            If Not rng Is Nothing Then
                FindAddress = rng.Address
            Do
                rng.Interior.ColorIndex = 6
                Set rng = .FindNext(rng)
            Loop While Not rng Is Nothing _
               And rng.Address <> FindAddress
               
            End If
        End With
    End If
    Set rng = Nothing
    
End Sub


'5-2 使用Like运算符进行模式匹配查找
Sub RngLike()
    
    Dim rng As Range
    Dim r As Integer
    r = 1
    Sheet1.Range("a:a").ClearContents
    For Each rng In Sheet2.Range("a1:a40")
        If rng.Text Like "*A*" Then
            Cells(r, 1) = rng.Text
            r = r + 1
        End If
    Next
    Set rng = Nothing
End Sub

'范例6 替换单元格内字符串
Sub Replaxement()

    Range("a:a").Replace _
    what:="市", replacement:="区", _
    lookat:=xlPart, Searchorder:=xlByRows, _
    MatchCase:=True
    
End Sub

'7-1 复制单元格区域
Sub RngCopy()
    Sheet1.Range("A1:G7").Copy Sheet2.Range("A1")

End Sub


Sub Copyalltheforms()

    Dim i As Integer
    Sheet1.Range("a1:g7").Copy
    With Sheet3.Range("a1")
        .PasteSpecial xlPasteAll
        .PasteSpecial xlPasteColumnWidths
    End With
    Application.CutCopyMode = False
    For i = 1 To 7
        Sheet3.Rows(i).RowHeight = Sheet1.Rows(i).RowHeight
    Next

End Sub


'7-2 仅复制数值到另一区域
Sub CopyValue()

    Sheet1.Range("a1:g7").Copy
    Sheet2.Range("a1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
End Sub


'除了copy,还有直接赋值法
Sub GetValueResize()
    With Sheet1.Range("a1").CurrentRegion
        Sheet3.Range("a1").Resize(.Rows.Count, .Columns.Count).Value = .Value
    End With

End Sub


'范例8 禁用单元格拖放功能
Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    If Target.Column = 1 Then
        Application.CellDragAndDrop = False
    Else
        Application.CellDragAndDrop = True
    End If
    
End Sub




'范例9 设置单元格格式

Sub CellFont()

    With Range("a1").Font
        .Name = "华文彩云"
        .FontStyle = "Bold"
        .Size = 22
        .ColorIndex = 3
        .Underline = 2
    End With
    
End Sub


'9-2 设置单元格内部格式
Sub CellInternalFormat()
    With Range("a4").Interior
        .ColorIndex = 3
        .Pattern = xlPatternGrid
        .PatternColorIndex = 6
    End With

End Sub


'9-3 为单元格区域添加边框

Sub CellBorder()
    Dim rng As Range
    Set rng = Range("b2:e8")
    
    '横
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .Weight = xlThin
        .ColorIndex = xlColorIndexAutomatic
    End With
    '竖
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlColorIndexAutomatic
    End With
    
    '周围
    rng.BorderAround xlContinuous, xlMedium, xlColorIndexAutomatic
    Set rng = Nothing
    

End Sub

'范例10 单元格的数据有效性

'10-1 添加数据有效性
Sub AddValidation()

    With Range("a1:a10").Validation
        .Delete
        .Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, _
            Formula1:="1,2,3,4,5,6,7,8"
        .ErrorMessage = "只能输入1-8的数值,请重新输入!"
    End With
    
End Sub

'10-2 判断是否存在数据有效性

Sub ErrValidation()
    On Error GoTo line
    If Range("a1").Validation.Type >= 0 Then
        MsgBox "有数据有效性！"
        Exit Sub
    End If
line:
    MsgBox "没有数据有效性!"

End Sub



'范例11 单元格中的公式  sheet5
'11-1 在单元格中写入公式
Sub rngFormula()

    Dim r As Integer
    r = Cells(Rows.Count, 1).End(3).Row
    Range("c2").Formula = "=a2*b2"
    Range("c2").Copy Range("c3:c" & r)
    Range("a" & r + 1) = "合计"
    Range("c" & r + 1).Formula = "=sum(c2:c" & r & ")"
    
End Sub


Sub rngFormulaRC()

    Dim r As Integer
    r = Cells(Rows.Count, 1).End(3).Row
    Range("c2:c" & r).Formula2R1C1 = "=rc[-2]*rc[-1]"
    Range("a" & r + 1) = "合计"
    Range("c" & r + 1).FormulaR1C1 = "SUM(R[-" & r - 1 & "]C:R[-1]C)"
End Sub


'11-2 判断单元格是否包含公式

Sub rngIsHasFormula()
    Select Case Selection.HasFormula
        Case True
            MsgBox "单元格包含公式！"
        Case False
            MsgBox "单元格没有公式!"
        Case Else
            MsgBox "公式区域：" & Selection.SpecialCells(-4123, 23).Address(0, 0)
    End Select
    
End Sub


'11-3 判断单元格公式是否存在错误
Sub CellFormulaIsWrong()
    
    If IsError(Range("c6").Value) = True Then
        MsgBox "a1单元格错误类型为:" & Range("c6").Text
    Else
        MsgBox "a1单元格公式结果为:" & Range("c6").Value
    End If

    'range的text属性与value属性的区别
    MsgBox "text:" & Range("a1").Text & vbCrLf & "value:" & Range("a1").Value
    
End Sub



'11-4 取得公式的引用单元格
Sub RngPrecedent()

    Dim rng As Range
    Set rng = Sheet5.Range("c10").Precedents
    MsgBox "公式所引用的单元格是 ： " & rng.Address
    Set rng = Nothing
    
End Sub


'11-5 将公式转换为数值
Sub SpecialPaste()
    With Range("c1:c9")
        .Copy
        .PasteSpecial xlPasteValues
    End With
    Application.CutCopyMode = False
    

'使用range对象的Value属性, 将函数公式结果转为数值
Range("a1:a10").Value = Range("a1:a10").Value

End Sub


'范例12 为单元格添加批注
Sub AddComment()

    With Range("a1")
        If Not .Comment Is Nothing Then .Comment.Delete
            .AddComment Text:=Date & vbCrLf & .Text
            .Comment.Visible = True
    End With
    
End Sub

'范例13 合并单元格操作
Sub IsMergeCell()
    
    If Range("a1").MergeCells Then
        MsgBox "合并单元格", vbInformation
    Else
        MsgBox "非合并单元格", vbInformation
    End If

End Sub



Sub IsMergeCells()
    If IsNull(Range("a1:d10").MergeCells) Then
        MsgBox "包含合并单元格", vbInformation
    Else
        MsgBox "没有包含合并单元格", vbInformation
    End If

End Sub


'13-2 合并单元格时连接每个单元格的文本

Sub MergeCells()

    Dim MergeStr As String
    Dim MergeRng As Range
    Dim rng As Range
    
    Set MergeRng = Range("a1:b2")
    For Each rng In MergeRng
        MergeStr = MergeStr & rng & " "
    Next
    Application.DisplayAlerts = False
    MergeRng.Merge (True)            'True 则将指定区域内的每一行合并为一个合并单元格，默认为False
    MergeRng.Value = MergeStr
    Application.DisplayAlerts = True
    Set MergeRng = Nothing
    Set rng = Nothing
    

End Sub


'13-3 合并内容相同的连续单元格

Sub MergeLinkedCell()

    Dim r As Integer
    Dim i As Integer
    Application.DisplayAlerts = False
    With Sheet5
        r = .Cells(Rows.Count, 1).End(3).Row
        For i = r To 2 Step -1
            If .Cells(i, 1).Value = .Cells(i - 1, 1).Value Then
                .Range(.Cells(i - 1, 1), .Cells(i, 1)).Merge
            End If
        Next
    End With
End Sub


'13-4 取消合并单元格时在每个单元格中保留内容

Sub CancelMergeCells()

    Dim r As Integer
    Dim MergeStr As String
    Dim MergeCot As Integer
    Dim i As Integer
    
    With Sheet5
        r = .Cells(.Rows.Count, 1).End(3).Row
        For i = 2 To r
            MergeStr = .Cells(i, 1).Value
            MergeCot = .Cells(i, 1).MergeArea.Count
            .Cells(i, 1).UnMerge
            .Range(.Cells(i, 1), .Cells(i + MergeCot - 1, 1)).Value = MergeStr
            i = i + MergeCot - 1
        Next
        .Range("a1:a" & r).Borders.LineStyle = xlContinuous
    End With
End Sub


''范例14 高亮显示选定单元格区域  sheet5
'Private Sub worksheet_selectionchange(ByVal Target As Range)
'
'    Cells.Interior.ColorIndex = xlColorIndexNone
'    Target.Interior.ColorIndex = Int(56 * Rnd() + 1)
'
'End Sub


''范例14-1 高亮显示选定单元格所在的行与列  sheet5
'Private Sub worksheet_selectionchange2(ByVal Target As Range)
'    Dim rng As Range
'    Cells.Interior.ColorIndex = xlColorIndexNone
'    Set rng = Application.Union(Target.EntireColumn, Target.EntireRow)
'    rng.Interior.ColorIndex = Int(56 * Rnd() + 1)
'    Set rng = Nothing
'End Sub



'范例 15 双击被保护单元格时不弹出提示消息框  Sheet5

'Private Sub worksheet_beforedoubleclick(ByVal Target As Range, Cancel As Boolean)
'   If Target.Locked = True Then
'        MsgBox "此单元格已保护,不能编辑"
'        Cancel = True
'    End If
'
'End Sub


'范例 16 单元格录入数据后的自动保护  sheet6
'Private Sub worksheet_selectionchange(ByVal Target As Range)
'    Dim msg As Byte
'    With Target
'        If Not Application.Intersect(Target, Range("a2:f6")) Is Nothing Then
'            If .Count > 1 Then
'            Range("a1").Select
'            Exit Sub
'        End If
'        ActiveSheet.Unprotect
'            If Len(Trim(.Value)) > 0 Then
'                msg = MsgBox("当前单元格已录入数据，是否修改? ", 32 + 4)
'                .Locked = IIf(msg = 6, False, True)
'            End If
'        ActiveSheet.Protect
'        ActiveSheet.EnableSelection = 0
'        End If
'    End With
'End Sub



'范例17 Target参数的使用方法
'
'Private Sub workbook_sheetselectionchange(ByVal Sh As Object, ByVal Target As Range)
'    Select Case Target.Address(0, 0)
'        Case "a1"
'            Sh.Unprotect
'        Case "a2"
'            Sh.Protect
'        Case Else
'    End Select
'End Sub

'17-3 使用intersect属性  sheet9
'使用Intersect属性可以很方便地指定一个或多个区域范围
'Private Sub WorkSheet_SelectionChange(ByVal Target As Range)
'
'    If Not Application.Intersect(Target, Union(Range("a1:b10"), Range("E1:F10"))) Is Nothing Then
'        If Target.Count = 1 Then
'            MsgBox "你选择了" & Target.Address(0, 0) & "单元格"
'        End If
'    End If
'
'End Sub