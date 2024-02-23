'container status2021.xlsx

'输入trip号码自动找出相关体积，客户信息

Private Sub Worksheet_Change(ByVal Target As Range)
    '如果更改的单元各不是B列第3行以下的单元格或更改的单元格个数大于1时退出程序
    If Application.Intersect(Target, Range("B2:B65536")) Is Nothing Or Target.Count > 1 Then
        Exit Sub
    End If
    
    Dim i As Integer
    i = 2
    Do While Sheet7.Cells(i, "a").Value <> ""   ' 在参照表中循环
        '判断录入的字母与参照表的字母是否相符
        If Mid(Target.Value, 3, 5) = Sheet7.Cells(i, "a").Value Then
                Target.Offset(0, -1).Value = Date
                Target.Offset(0, 1).Value = Sheet7.Cells(i, "a").Offset(0, 3).Value
            GoTo 100
        End If
        i = i + 1
    Loop
    
100
    Dim x As Integer
    x = 2
    Do While Sheet1.Cells(x, "a").Value <> ""   ' 在参照表中循环
        '判断录入的字母与参照表的字母是否相符
        If Mid(Target.Value, 3, 5) = Sheet1.Cells(x, "a").Value Then
                Target.Offset(0, 5).Value = Sheet1.Cells(x, "a").Offset(0, 6).Value
                Target.Offset(0, 16).Value = Sheet1.Cells(x, "a").Offset(0, 7).Value
            Exit Sub
        End If
        x = x + 1
    Loop
    
End Sub




