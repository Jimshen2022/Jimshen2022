Sub ConsolidateSheets()
    'EH技术论坛公众号VBA编程学习与实践
    Dim Sht As Worksheet
    Dim r, k&, i&
    ReDim r(1 To 1)
    For Each Sht In Worksheets
        '遍历工作表
        If Sht.Name <> ActiveSheet.Name Then
            k = k + 1
            ReDim Preserve r(1 To k) '动态设置数组大小
            r(k) = Sht.Name & "!" & Sht.UsedRange.Address(ReferenceStyle:=xlR1C1)
            '数据区域地址以r1c1形式装入数组r
        End If
    Next
    Cells.ClearContents '清除当前表数据
    Range("a1").Consolidate Sources:=r, Function :=xlSum, toprow:=True, leftcolumn:=True
        'Consolidate合并计算语句，基于行列汇总，求和形式。
        MsgBox "合并计算OK"
    End Sub
