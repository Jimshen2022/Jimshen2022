Sub GetShtByVba()
    Dim sht As Worksheet, k As Long
    Application.ScreenUpdating = False
    k = 1
    Range("a:b").Clear '清空数据
    Range("a:a").NumberFormat = "@" '设置文本格式
    For Each sht In Worksheets '遍历工作表取表名
        k = k + 1
        Cells(k, 1) = sht.Name
    Next
    Range("a1:b1") = Array("工作表名", "是否删除")
    Application.ScreenUpdating = True
End Sub