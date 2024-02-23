以下代码可以删除当前工作簿的首个工作表 。
Sub DelSht()
    Application.DisplayAlerts = False
    Worksheets(1).Delete
    Application.DisplayAlerts = True
End Sub


以下代码可以删除当前工作簿 "全部" 的工作表 ……。
Sub DelShtAll()
    Dim sht As Worksheet
    Application.DisplayAlerts = False
    For Each sht In Sheets '集合遍历
        If sht.Name <> ActiveSheet.Name Then
            sht.Delete '如果sht的名字不等于当前工作表则删除
        End If
    Next
    Application.DisplayAlerts = True
End Sub

删除指定名单工作表


Sub DelShtByCustom()
    Dim sht As Worksheet, rngData As Range, c As Range
    Dim d As Object, y As Long
    Dim strName As String, strErr As String
    If ActiveWorkbook.ProtectStructure = True Then
        MsgBox "工作簿有保护，需要先撤销保护再运行代码"
        Exit Sub
    End If
    On Error Resume Next '使程序忽略错误继续运行
    Set rngData = Application.InputBox("请选择需要删除的工作表名单区域",  _
            Title:="公众号Excel星球",  _
            Default :=Selection.Address,  _
            Type :=8)
        Set rngData = Intersect(rngData, rngData.Parent.UsedRange)
        If rngData Is Nothing Then
            MsgBox "未选择有效数据区域。"
            Exit Sub
        End If
        Set d = CreateObject("scripting.dictionary") '后期字典
        For Each sht In Sheets '遍历工作表名存入字典
            strName = sht.Name
            d(strName) = ""
        Next
        With Application '取消屏幕刷新、信息警告等
             .ScreenUpdating = False
             .DisplayAlerts = False
             .Calculation = xlCalculationManual
        End With
        For Each c In rngData '遍历名单区域
            strName = c.Value
            If Len(strName) Then '如果名字非空
                If d.exists(strName) Then '如果字典中存在删除表名
                    If Sheets.Count > 1 Then '判断工作表个数是否可删
                        Sheets(strName).Delete '删除工作表
                        y = y + 1 '累加个数
                    Else
                        MsgBox "系统要求工作表必须保留至少一张，因此" &  _
                                strName & "未能删除。"
                    End If
                Else '如果不存在删除表名
                    strErr = strErr & "," & strName '合并不存在的表名
                End If
            End If
        Next
        With Application '恢复屏幕刷新、信息警告等
             .ScreenUpdating = True
             .DisplayAlerts = True
             .Calculation = xlCalculationAutomatic
        End With
        If strErr <> "" Then
            MsgBox "以下名称工作簿中不存在工作表，未能删除：" & VbCrLf _
                     & Mid(strErr, 2)
        Else
            MsgBox "处理完成。"
        End If
        Set d = Nothing
    End Sub