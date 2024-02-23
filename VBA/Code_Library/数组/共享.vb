方法
Sub 共享()
    With ActiveWorkbook
         .KeepChangeHistory = True
         .ChangeHistoryDuration = 30
    End With
    ActiveWorkbook.SaveAs Filename:= _
            "D:\共享数据.xls", AccessMode:=xlShared
    ActiveWorkbook.Close
    MsgBox "共享设置已经完成" & Chr(10) & "共享后工作薄另存在D:\‘共享数据’，点击确定后将自动打开直接进行操作！", 64, "共享设置"
    Workbooks.Open Filename:= _
            "D:\共享数据.xls"
    ActiveWorkbook.Save
    End
End Sub