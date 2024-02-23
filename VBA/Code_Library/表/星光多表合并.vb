Sub ByVBA()
    Dim strFindSht As String, sht As Worksheet, k As Long
    Dim m As Long, aData
    Application.ScreenUpdating = False
    strFindSht = ",一部门,二部门,三部门,四部门,后勤部,"
    Worksheets("VBA").Select
    Cells.ClearContents
    Cells.NumberFormatLocal = "@" '设置文本格式，避免工号变形
    k = 1
    For Each sht In Worksheets '遍历工作表
        If InStr(strFindSht, "," & sht.Name & ",") Then '判断工作表名称是否符合条件
            k = k + m '放置数据的开始行
            If k = 1 Then
                aData = sht.UsedRange
            Else
                aData = sht.UsedRange.Offset(1) '扣掉标题行
            End If
            m = UBound(aData) + (k > 1) '注意在VBA中True等于-1
            Cells(k, 1).Resize(m) = sht.Name '工作表名称
            Cells(k, 2).Resize(m, UBound(aData, 2)) = aData '数据
        End If
    Next
    Cells(1, 1) = "工作表名称"
    Application.ScreenUpdating = True
End Sub
