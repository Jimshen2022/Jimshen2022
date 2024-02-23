
'5.5.3 让文件每隔一分钟自动保存一次


Private Sub Workbook_Open()
 Call otime    '打开工作簿后自动运行otime过程
End Sub



Sub otime()
    '一分钟自动运行Wbsave过程
    Application.OnTime Now() + TimeValue("00:00:30"), "wbsave"
    
End Sub


Sub wbsave()
    ThisWorkbook.Save
    Call otime
    
End Sub


