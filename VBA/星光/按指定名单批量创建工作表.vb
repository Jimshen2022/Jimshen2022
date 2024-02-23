Sub NewShtBySelection()
    Dim shtAct As Worksheet
    Dim rngData As Range, c As Range
    Dim strName As String
    Dim n As Long, y As Long, strErr As String
    If ActiveWorkbook.ProtectStructure = True Then
        MsgBox "工作簿有保护，无法新建工作表，请先撤除保护。"
        Exit Sub
    End If
    On Error Resume Next '忽略程序错误继续运行
    Set rngData = Application.InputBox("请选择新建工作表名称来源。",  _
            Title:="提示",  _
            Default :=Selection.Address,  _
            Type :=8) '用户选择名称来源区域
        Set rngData = Intersect(rngData, rngData.Parent.UsedRange)
        '交集运算，避免用户选择整列数据造成运算量虚大或选择区域空白
        If rngData Is Nothing Then '如果用户关闭了对话框，或选择区域空白，则退出程序
            MsgBox "未选择有效区域。"
            Exit Sub
        End If
        Set shtAct = ActiveSheet '当前工作表，操作完成后界面回到这里
        With Application '取消屏幕刷新、信息警告、公式重算等
             .ScreenUpdating = False
             .DisplayAlerts = False
             .AskToUpdateLinks = False
             .Calculation = xlCalculationManual
        End With
        For Each c In rngData '遍历名单
            strName = c.Value '工作表名称
            If Len(strName) Then '如果工作表名称非空
                Err.Clear '清除错误
                Worksheets.Add after:=Sheets(Sheets.Count) '新建工作表
                ActiveSheet.Name = strName '命名
                If Err.Number Then '如果存在错误，说明有重名或工作表名称不规范
                    ActiveSheet.Delete '删除新建工作表
                    n = n + 1 '记录问题名称数量
                    strErr = strErr & "," & strName '记录名称
                Else
                    y = y + 1 '记录正确创建工作表的数量
                End If
            End If
        Next
        shtAct.Select
        With Application
             .ScreenUpdating = True
             .DisplayAlerts = True
             .AskToUpdateLinks = True
             .Calculation = xlCalculationAutomatic
        End With
        If n Then
            MsgBox "有" & n & "张工作表创建失败，原因是工作表重名或格式错误。" &  _
                    "名单如下：" & VbCrLf &  _
                    Mid(strErr, 2)
        ElseIf y Then
            MsgBox "创建完成。"
        End If
    End Sub