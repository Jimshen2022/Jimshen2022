

Sub GetFilesDataByNUM()
    Dim aFileName(), strPath As String
    Dim i As Long, x As Long, k As Long, intTitCount
    Dim wb As Workbook, sht As Worksheet, shtSum As Worksheet
    Dim rngData As Range
    Dim intLastRow As Long, intFirstRow As Long
    Dim aData, aSource
'    On Error Resume Next
    strPath = "\\wanvogfurniture.com\wvf-dfs\Kunshan\Dept\Warehouse Operations\Internal\Daily Ending report-shipping\Shipping\PPH\2022\"   '用户选择路径
    If strPath = "" Then Exit Sub
    intTitCount = Sheet3.Range("b1").Value   '用户设置标题行数
    If intTitCount = "错误" Then Exit Sub
    aFileName = GetWbFullNames(strPath) '获取文件名单
    Call disAppSet '取消屏幕刷新
    Call CreateShtSum '创建汇总数据的工作表
    Set shtSum = Worksheets("每天资料汇总")
    intFirstRow = 1
    For i = 1 To UBound(aFileName) '遍历文件
        Set wb = Workbooks.Open(aFileName(i))
        For Each sht In wb.Worksheets '遍历工作表
             If sht.Name <> "Billable & CPW (2)" Then
                Exit For
             Else
                 Set rngData = sht.Range("a1:x12")
                    If IsEmpty(rngData) = False Then '如果工作表非空
                        k = k + 1
                        '数据来源的工作簿、工作表等信息
                        aSource = Array(wb.Name, sht.Name, sht.Index, sht.range("b1").value)
                        If k = 1 Then
                            aData = rngData.Value
                            '根据首张工作表，设置可能有的文本值格式
                            Call DataFormat(aData, shtSum)
                        Else
                            aData = rngData.Offset(intTitCount).Value
                        End If
                        With shtSum '数据写入工作表
                            .Cells(intFirstRow, 5).Resize( _
                                    UBound(aData), UBound(aData, 2)) = aData
                            intLastRow = GetLastRow(shtSum) '结束行
                            .Range(.Cells(intFirstRow, 1), .Cells(intLastRow, 4)) _
                                    .Value = aSource '来源信息写入工作表
                            intFirstRow = intLastRow + 1
                        End With
                     End If
                End If
        Next
        wb.Close False
    Next
    shtSum.Select
    Range("a1:d1") = Array("工作簿名称", "工作表名称", "工作表索引","报表日期")
    Cells.EntireColumn.AutoFit
    Call reAppSet
    If Err.Number Then
        MsgBox Err.Description
    Else
        MsgBox "汇总完成。"
    End If
End Sub

'用户选择文件夹路径
Function getStrPath() As String
    Dim strPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show Then
            strPath = .SelectedItems(1)
        Else '如用户为选中文件夹则退出
            Exit Function
        End If
    End With
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    getStrPath = strPath
End Function

'获取用户输入的标题行数
Function getTitCount()
    Dim intTitCount
    intTitCount = InputBox("请输入标题行的行数", _
                        Title:="公众号Excel星球", _
                        Default:=1)
    If StrPtr(intTitCount) = False Then
        getTitCount = "错误"
        Exit Function
    End If
    If IsNumeric(intTitCount) = False Then
        MsgBox "标题行的行数只能输入数字。"
        getTitCount = "错误"
        Exit Function
    End If
    If intTitCount < 0 Then
        MsgBox "标题行数不能为负数。"
        getTitCount = "错误"
        Exit Function
    End If
    getTitCount = intTitCount
End Function

'判断是否文本格式，由前10行决定
Sub DataFormat(ByRef aData As Variant, shtSum As Worksheet)
    Dim i As Long, j As Long
    Dim vnt, strADS
    For j = 1 To UBound(aData, 2) '遍历列
        For i = 1 To UBound(aData) '遍历前10行
            If i > 10 Then Exit For
            vnt = aData(i, j)
            If IsNumeric(vnt) Then '是否数值
                If VarType(aData(i, j)) = 8 Then '是否文本
                    strADS = strADS & "," & Cells(1, j + 3).Address
                    Exit For
                End If
            End If
        Next
    Next
    strADS = Mid(strADS, 2) '需要设置文本格式的单元格地址
    If Len(strADS) Then
        shtSum.Range(strADS).EntireColumn.NumberFormat = "@"
    End If
End Sub

'获取文件名名单
Function GetWbFullNames(strPath As String)
    Dim strName As String, strTemp As String
    Dim aRes(), k As Long
    strName = Dir(strPath & "*.*")
    Do While strName <> ""
        strTemp = Right(strName, 4)
        If strTemp Like "*xls*" Or strTemp Like "*csv*" Then
            k = k + 1
            ReDim Preserve aRes(1 To k)
            aRes(k) = strPath & strName
        End If
        strName = Dir()
    Loop
    GetWbFullNames = aRes
End Function

'创建汇总表
Sub CreateShtSum()
    Dim sht As Worksheet
    For Each sht In Worksheets
        If sht.Name = "每天资料汇总" Then sht.Cells.Clear
    Next
'    Worksheets.Add , Sheets(1)
'    ActiveSheet.Name = "每天资料汇总"
End Sub

'查询有效数据最大行
Function GetLastRow(shtData As Worksheet)
    GetLastRow = shtData.Cells.Find("*", _
        LookIn:=xlFormulas, SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious).Row
End Function

Sub disAppSet() '撤销屏幕刷新
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
        .AskToUpdateLinks = False
        .Calculation = xlCalculationManual
    End With
End Sub

Sub reAppSet() '恢复屏幕刷新等
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
        .AskToUpdateLinks = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub


