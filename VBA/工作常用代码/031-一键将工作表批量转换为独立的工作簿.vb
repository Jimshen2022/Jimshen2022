
Sub EachShtToWorkbook()
    Dim Sht As Worksheet, strPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
    '选择保存工作薄的文件路径
        If .Show Then strPath = .SelectedItems(1) Else Exit Sub
        '读取选择的文件路径,如果用户未选取文件路径则退出程序
    End With
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    Application.DisplayAlerts = False
    '取消显示系统警告消息，避免重名工作簿无法保存。当有重名工作簿时，会直接覆盖保存
    Application.ScreenUpdating = False
    For Each Sht In Worksheets
        Sht.Copy     '复制工作表，工作表单纯复制后，会成为活动工作表
        With ActiveWorkbook
            .SaveAs strPath & Sht.Name, xlWorkbookDefault
            '保存活动工作簿到指定路径下，以当前系统默认文件格式
            .Close True  '关闭工作簿并保存
        End With
    Next
    MsgBox "处理完成. ", , "提醒"
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
        
        