Sub CollectData()
    
    Dim Sht As Worksheet, rng As Range, k&, n&
    
    Application.ScreenUpdating = False
    
    n = Val(InputBox("Please enter title rows", "Notice"))
    
    If n < 0 Then MsgBox " the title row must greater than 0.", 64, "Notice": Exit Sub
    '取得用户输入的标题行数，如果为负数，退出程序
    
    Cells.ClearContents '清空当前表数据
    
    For Each Sht In Worksheets
        If Sht.Name <> ActiveSheet.Name Then '如果工作表名称不等于当前表名则进行汇总动作……
            Set rng = Sht.UsedRange '定义rng为表格已用区域
            k = k + 1 '累计表的个数
            If k = 1 Then '如果是首个表格，则把标题行一起复制到汇总表
                rng.Copy
                Cells(1, 1).PasteSpecial Paste:=xlPasteValues
            Else '否则，扣除标题行后再复制黏贴到总表
                rng.Offset(n).Copy
                Cells(ActiveSheet.UsedRange.Rows.Count + 1, 1).PasteSpecial Paste:=xlPasteValues
            End If
        End If
        
    Next
    Cells(1, 1).Activate
    Application.ScreenUpdating = True
    
    
End Sub