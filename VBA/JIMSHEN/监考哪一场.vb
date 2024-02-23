'5.5.2 监考哪一场
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Range("a2:t36").Interior.ColorIndex = xlNone   '清除单元格里原有底纹颜色
    '当选中的单元格个数大于1时，重新给Target赋值
    If Target.Count > 1 Then
        Set Target = Target.Cells(1)
    End If
    
    '当选中的单元格不包含指定区域的单元格时，退出程序
    If Application.Intersect(Target, Range("a2:t36")) Is Nothing Then
        Exit Sub
    End If
    
    Dim rng As Range
    For Each rng In Range("a2:t36")
        If rng.Value = Target.Value Then
            rng.Interior.ColorIndex = 39
        End If
        
    Next

End Sub




'高亮显示工作表中选中单元格所在的行和列

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Range("a2:t36").Interior.ColorIndex = xlNone
    If Target.Count > 1 Then
        Set Target = Target.Cells(1)
    End If
    
    '当选中的单元格不包含指定的单元格时，退出程序
    If Application.Intersect(Target, Range("a2:t36")) Is Nothing Then
        Exit Sub
    End If
    If Application.Intersect(Target, Range("a2:t36")) Is Nothing Then
        Exit Sub
    End If
    '添加底纹颜色
    Range(Cells(Target.Row, "a"), Cells(Target.Row, "t")).Interior.ColorIndex = 39
    Range(Cells(2, Target.Column), Cells(36, Target.Column)).Interior.ColorIndex = 39
    
End Sub