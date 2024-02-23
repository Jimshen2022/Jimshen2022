Sub DelShtByVba()
    Dim sht As Worksheet, i As Long, r
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error Resume Next
    r = Range("a1").CurrentRegion '数据装入数组r
    For i = 2 To UBound(r) '遍历并删除工作表
        If r(i, 2) = "删除" Then Worksheets(CStr(r(i, 1))).Delete
    Next
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub