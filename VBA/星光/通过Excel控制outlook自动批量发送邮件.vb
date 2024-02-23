Sub SendMailEnvelope()
    Dim avntWage As Variant
    Dim i As Long
    Dim strText As String
    Dim objAttach As Object
    Dim strPath As String
    With Application
         .ScreenUpdating = False
         .EnableEvents = False
    End With
    strPath = ThisWorkbook.Path & "\关于企业调整职工工资的通知.docx"
    '------------邮件发送附件的路径
    avntWage = Sheets("工资表").[a1].CurrentRegion
    '------------工资表的数据装入数组
    For i = 2 To UBound(avntWage)
        [a2:i2] = Application.Index(avntWage, i)
        '------------工资条数据放入a2:i2区域
        [b1:i2].Select
        '------------选中b1:i2作为邮件正文的表格内容
        ActiveWorkbook.EnvelopeVisible = True
        '------------MailEnvelope可见
        With ActiveSheet.MailEnvelope
            strText = avntWage(i, 2) & "您好：" & VbCrLf & "以下是您" &  _
                    avntWage(i, 3) & "月份工资明细，请查收！"
             .Introduction = strText
            '------------邮件正文内容
            With  .Item
                 . To =avntWage(i, 1)
                '------------收件人
                 .CC = "treasurer@gmail.com"
                '------------抄送人
                 .Subject = avntWage(i, 3) & "月份工资明细"
                '------------主题
                Set objAttach =  .Attachments
                Do While objAttach.Count > 0
                    '------------Do While语句删除可能存在的旧附件
                    objAttach.Remove 1
                    MsgBox objAttach.Count
                Loop
                 .Attachments.Add strPath
                '------------添加新附件
                 .send
                '------------发送邮件
            End With
        End With
    Next i
    ActiveWorkbook.EnvelopeVisible = False
    With Application
         .ScreenUpdating = True
         .EnableEvents = True
    End With
    Set objAttach = Nothing
End Sub