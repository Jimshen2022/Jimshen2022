'范例35 复制自动筛选后的数据区域
Sub CopyFilter()
    Sheet2.Cells.Clear
    With Sheet18
        If .FilterMode Then
            .AutoFilter.Range.SpecialCells(12).Copy Sheet2.Cells(1, 1)
        End If
    End With
End Sub
