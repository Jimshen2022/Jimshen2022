'每个工作簿包含数量不一 、名称相同的工作表 ，在汇总数据时 ，需要按工作表名称分别汇总 。比如 ，名为财务部的工作表单独汇总成一张工作表 ，名为销售部的也单独汇总成一张工作表 …


Sub GetEachShtData()
    Dim i As Long, intLastRow As Long
    Dim shtSum As Worksheet, shtAct As Worksheet, shtData As Worksheet
    Dim aFileName, wb As Workbook, d As Object
    Dim strFileName As String, strPath As String, strShtName As String
    On Error Resume Next
    strPath = getStrPath() '用户选择路径
    If strPath = "" Then Exit Sub
    aFileName = GetWbFullNames(strPath) '获取文件名单
    If IsArray(aFileName) = False Then Exit Sub
    Call disAppSet '取消屏幕刷新等
    Call delsht '调用删除工作表过程
    Set d = CreateObject("scripting.dictionary")
    Set shtAct = ActiveSheet '当前工作表
    Set wb = ThisWorkbook '代码所在工作簿
    For i = 1 To UBound(aFileName) '遍历工作簿
        With Workbooks.Open(aFileName(i), False) '打开工作簿不更新链接
            For Each shtData In  .Worksheets
                If shtData.FilterMode = True Then shtData.Cells.AutoFilter '取消筛选
                strShtName = shtData.Name '工作表名称
                If Not d.exists(strShtName) Then
                    d(strShtName) = "" '工作表移动到代码所在工作簿
                    shtData.Copy after:=wb.Worksheets(wb.Sheets.Count)
                Else
                    Set shtSum = wb.Worksheets(strShtName)
                    intLastRow = GetLastRow(shtSum) + 1 '最后存在数据的行
                    shtData.UsedRange.Copy shtSum.Cells(intLastRow, 1) '复制粘贴
                End If
            Next
             .Close False '关闭不保存
        End With
    Next
    Call reAppSet '恢复系统设置
    Set d = Nothing
    shtAct.Select
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
        If  .Show Then
            strPath =  .SelectedItems(1)
        Else '如用户为选中文件夹则退出
            Exit Function
        End If
    End With
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    getStrPath = strPath
End Function

'获取文件名名单
Function GetWbFullNames(strPath As String)
    Dim strShtName As String, strTemp As String
    Dim aRes(), k As Long
    k = 0
    strShtName = Dir(strPath & "*.*")
    Do While strShtName <> ""
        strTemp = Right(strShtName, 4)
        If strTemp Like "*xls*" Or strTemp Like "*csv*" Then
            k = k + 1
            ReDim Preserve aRes(1 To k)
            aRes(k) = strPath & strShtName
        End If
        strShtName = Dir()
    Loop
    GetWbFullNames = aRes
End Function

'查询有效数据最大行
Function GetLastRow(shtData As Worksheet)
    GetLastRow = shtData.Cells.Find("*",  _
            LookIn:=xlFormulas, SearchOrder:=xlByRows,  _
            SearchDirection:=xlPrevious).Row
End Function

Sub delsht()
    Dim sht As Worksheet
    For Each sht In ThisWorkbook.Worksheets
        If sht.Name <> ActiveSheet.Name Then sht.Delete
    Next
End Sub

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