'如何按指定模板批量创建工作表 ？

Sub NewShtByTemp()
    Dim shtAct As Worksheet, shtTemp As Worksheet
    Dim rngData As Range, strName As String, c As Range
    Dim n As Long, y As Long, strErr As String
    On Error Resume Next
    If ActiveWorkbook.ProtectStructure = True Then
        MsgBox "工作簿有保护，无法新建工作表，请先撤除保护。"
        Exit Sub
    End If
    Set rngData = Application.InputBox("请选择新建工作表名称来源。",  _
            Title:="公众号Excel星球",  _
            Default :=Selection.Address,  _
            Type :=8) '用户选择名称来源区域
        Set rngData = Intersect(rngData, rngData.Parent.UsedRange)
        '交集运算，避免用户选择整列数据造成运算量虚大或选择区域空白
        If rngData Is Nothing Then '如果用户关闭了对话框，或选择区域空白，则退出程序
            MsgBox "未选择有效区域。"
            Exit Sub
        End If
        Set shtTemp = Worksheets("模板")
        If Err.Number Then
            MsgBox "HI，没找到名为模板的工作表，请核实。"
            Exit Sub
        End If
        Set shtAct = ActiveSheet '当前工作表，操作完成后界面回到这里
        With Application '取消系统刷新、警告、链接、公式重算等
             .ScreenUpdating = False
             .DisplayAlerts = False
             .AskToUpdateLinks = False
             .Calculation = xlCalculationManual
        End With
        For Each c In rngData '遍历名单
            strName = c.Value '工作表名称
            If Len(strName) Then '如果工作表名称非空
                Sheets(strName).Delete '删除可能存在的旧表
                Err.Clear '清除错误记录
                shtTemp.Copy after:=Sheets(Sheets.Count) '复制一个模板表
                ActiveSheet.Name = strName '命名
                If Err.Number Then '如果存在错误，说明有重名或工作表名称不规范
                    ActiveSheet.Delete '删除已新建工作表
                    n = n + 1 '记录问题名称数量
                    strErr = strErr & "," & strName '记录名称
                Else
                    y = y + 1 '记录正确创建工作表的数量
                End If
            End If
        Next
        shtAct.Select
        With Application '恢复系统设定
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