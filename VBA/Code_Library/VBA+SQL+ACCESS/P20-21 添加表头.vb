Option Explicit


Private Sub 添加表头_Click()
    With ListView1
        '方法1：挨个添加，比较适合列数已知且不太多的情况
        '        .ColumnHeaders.Add 1, "xh", "学号", .Width / 7, lvwColumnLeft
        '        .ColumnHeaders.Add 2, "xm", "姓名", .Width / 7, lvwColumnLeft
        '        .ColumnHeaders.Add 3, "bj", "班级", .Width / 7, lvwColumnLeft
        '        .ColumnHeaders.Add 4, "yw", "语文", .Width / 7, lvwColumnLeft
        '        .ColumnHeaders.Add 5, "sx", "数学", .Width / 7, lvwColumnLeft
        '        .ColumnHeaders.Add 6, "yy", "英语", .Width / 7, lvwColumnLeft
        '        .ColumnHeaders.Add 7, "zf", "总分", .Width / 7, lvwColumnLeft
        
        '方法2：利用循环动态添加表头
         .ColumnHeaders.Clear
        Dim i As Integer '循环变量
        Dim col As Integer '用于记录列数
        col = Range("A1").End(xlToRight).Column '从a1开始向右获取最后一列列号
        For i = 1 To col
             .ColumnHeaders.Add i, , Cells(1, i),  .Width / col, lvwColumnLeft
        Next i
        '格式处理
         .Gridlines = True '显示表格线
         .FullRowSelect = True '支持整行选择
         .View = lvwReport '设置数据以报表形式显示
        
        
        '循环添加记录
         .ListItems.Clear
        Dim j As Integer, ITM As ListItem
        For i = 2 To Range("a1").End(4).Row
            Set ITM =  .ListItems.Add
            For j = 1 To Range("a1").End(2).Column - 1
                ITM.Text = Cells(i, 1)
                ITM.SubItems(j) = Cells(i, j + 1)
            Next j
        Next i
    End With
End Sub

'Private Sub 添加记录_Click()
'    Dim ITM As ListItem     '记录的每一行称为list item
'    Dim i As Integer
'    '手动添加列
''    For i = 2 To Range("a1").End(4).Row
''        Set ITM = ListView1.ListItems.Add()
''        ITM.Text = Cells(i, 1)
''        ITM.SubItems(1) = Cells(i, 2)
''        ITM.SubItems(2) = Cells(i, 3)
''        ITM.SubItems(3) = Cells(i, 4)
''        ITM.SubItems(4) = Cells(i, 5)
''        ITM.SubItems(5) = Cells(i, 6)
''        ITM.SubItems(6) = Cells(i, 7)
''        ITM.SubItems(7) = Cells(i, 8)
''    Next i
'
'    '循环添加列
'    Dim j%
'    With ListView1
'        For i = 2 To Range("a1").End(4).Row
'            Set ITM = .ListItems.Add
'            For j = 1 To Range("a1").End(2).Column - 1
'                ITM.Text = Cells(i, 1)
'                ITM.SubItems(j) = Cells(i, j + 1)
'            Next j
'        Next i
'    End With
'End Sub



