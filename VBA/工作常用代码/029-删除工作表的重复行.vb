'范例31 删除工作表的重复行

Sub DeleteRow()
    Dim r As Integer, i As Integer
    With Sheet15
        r = .Cells(.Rows.Count, 1).End(3).Row
        For i = r To 1 Step -1
            If WorksheetFunction.CountIf(.Columns(4), .Cells(i, 4)) > 1 Then
                .Rows(i).Delete
            End If
        Next
    End With

End Sub



'范例30 删除工作表中的空行

Sub DelBlankRow()
    Dim r As Long
    Dim i As Long
    r = Sheet15.UsedRange.Rows.Count
    For i = r To 1 Step -1
        If Rows(i).Find("*", , xlValues, , , 2) Is Nothing Then
        Rows(i).Delete
        End If
    Next

End Sub


'范例29 在工作表中一次插入多行
Sub InSertRow()
    Dim i As Integer
    For i = 1 To 3
        Sheet15.Rows(5).Insert
    Next
End Sub