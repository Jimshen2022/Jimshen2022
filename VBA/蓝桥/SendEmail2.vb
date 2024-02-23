'工具->引用->Microsoft Outlook 16.0 Object Library
'或者  Set Mail = CreateObject("Outlook.Application")
Sub SendEmail()
    Dim mail As Outlook.Application
    Set mail = CreateObject("Outlook.Application")
    Dim objMail As Outlook.MailItem
    Set objMail = mail.CreateItem(olMailItem)
    
    With objMail
        .Subject = "MIL UPH Invenotry Report"  '主题
        .To = 收件人(Sheet6.[A:A]) '收件人
        .CC = 收件人(Sheet6.[B:B])  '抄送
'        .BCC = 收件人(Sheet6.[C:C])    '密送
         附件添加 .Attachments, Sheet6.[D:D] '添加附件
     
        .BodyFormat = olFormatHTML                  '正文添加图片,尺寸控制如下备注
'        .HTMLBody = "<style> h4{font-family:Century Gothic; font-weight:normal;font-size:15} </style> <h4>Hi,&ensp;Jane, <br>&emsp;I'd like to share these VBA code with you, hope it is useful for you. any questions, please let me know. thanks </h4>"
        .HTMLBody = "<style> h4{font-family:Century Gothic; font-weight:normal;font-size:15} </style> <h4>F.Y.I</h4>"
        .HTMLBody = .HTMLBody & Range_to_Html(Sheet3.[A1:I30])  '正文 添加表格
        .Display
        .Send '执行发送
    End With

'如果想得到能控制尺寸的正文图片，需要用如下方法进行：
'.Attachments.Add "F:\图片汇总\PHOTO.png"
'.BodyFormat = olFormatHTML
'.HTMLBody = "<img src='cid:PHOTO.png' width='100' height='100'>"
'.Display

End Sub



Private Sub 附件添加(附件 As Outlook.Attachments, Rng As Range)
    Dim Rr As Integer
    Rr = 2
    While Rng.Cells(Rr, 1) <> ""
        附件.Add Rng.Cells(Rr, 1).Text
        Rr = Rr + 1
    Wend
End Sub

Private Function 收件人(Rng As Range) As String
    收件人 = ""
    Dim Rr As Integer
    Rr = 2
    While Rng.Cells(Rr, 1) <> ""
        收件人 = 收件人 & Rng.Cells(Rr, 1)
        If Rng.Cells(Rr + 1, 1) <> "" Then 收件人 = 收件人 & ";"
        Rr = Rr + 1
    Wend
End Function

'工具->引用-> Microsoft Scripting Runtime
'该函数，要求一个期望加载到邮件正文的区域，其返回的就是代表那个表区域的HTML代码.
Function Range_to_Html(Rng As Range) As String
    Dim PO As PublishObject
    Set PO = ThisWorkbook.PublishObjects.Add(xlSourceRange, "D:\Result.htm", Rng.Parent.Name, Rng.Address, xlHtmlStatic)
    PO.Publish True
    PO.Delete
    
    Dim FS As FileSystemObject
    Set FS = New FileSystemObject
    Dim TS As TextStream
    Set TS = FS.OpenTextFile("D:\Result.htm", ForReading, True, TristateUseDefault)
    Range_to_Html = TS.ReadAll
End Function

