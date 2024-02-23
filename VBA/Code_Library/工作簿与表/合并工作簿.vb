'快速合并工作簿

Sub CltSheets()

    'ExcelHome技术论坛公众号：VBA编程学习与实践，作者看见星光

    Dim P$, Bookn$, Book$, Keystr1, Keystr2, Shtname$, K&

    Dim Sht As Worksheet, Sh As Worksheet

    Application.ScreenUpdating = False

    Application.DisplayAlerts = False

    On Error Resume Next

    With Application.FileDialog(msoFileDialogFolderPicker)

        .AllowMultiSelect = False

        If .Show Then P = .SelectedItems(1) Else: Exit Sub

    End With

    If Right(P, 1) <> "\" Then P = P & "\"

    Keystr1 = InputBox("请输入工作簿名称所包含的关键词。" & vbCr & "关键词可以为空，如为空，则默认选择全部工作簿")

    If StrPtr(Keystr1) = 0 Then Exit Sub '如果用户点击了取消或关闭按钮，则退出程序

    Keystr2 = InputBox("请输入工作表名称所包含的关键词。" & vbCr & "关键词可以为空，如为空，则默认选择符合条件工作簿的全部工作表")

    If StrPtr(Keystr2) = 0 Then Exit Sub

    Set Sh = ActiveSheet '当前工作表，赋值变量,代码运行完毕后，回到此表

    Bookn = Dir(P & "*.xls*")

    Do While Bookn <> ""

        If Bookn = ThisWorkbook.Name Then

            MsgBox "注意：指定文件夹中存在和当前表格重名的工作簿！！" & vbCr & "该工作簿无法打开，工作表无法复制。"

            '当出现重名工作簿时，提醒用户。

        Else

            If InStr(1, Bookn, Keystr1, vbTextCompare) Then

            '工作簿名称是否包含关键词，关键词不区分大小写

                With GetObject(P & Bookn)

                    For Each Sht In .Worksheets

                        If InStr(1, Sht.Name, Keystr2, vbTextCompare) Then

                        '工作表名称是否包含关键词，关键词不区分大小写

                            If Application.CountIf(Sht.UsedRange, "<>") Then

                            '如果表格存在数据区域

                                Shtname = Split(Bookn, ".xls")(0) & "-" & Sht.Name

                                '复制来的工作表以"工作簿-工作表"形式起名。

                                ThisWorkbook.Sheets(Shtname).Delete

                                '如果已存在相关表名，则删除

                                Sht.Copy after:=ThisWorkbook.Worksheets(Sheets.Count)

                                K = K + 1

                                '复制Sht到代码所在工作簿所有工作表的后面，并累计个数

                                ActiveSheet.Name = Shtname

                                '工作表命名。

                            End If

                        End If

                    Next

                    .Close False '关闭工作簿

                End With

            End If

        End If

        Bookn = Dir '下一个符合条件的文件

    Loop

    Sh.Select '回到初始工作表

    MsgBox "工作表收集完毕，共收集：" & K & "个"

    Application.ScreenUpdating = True

    Application.DisplayAlerts = True

End Sub
    
    


