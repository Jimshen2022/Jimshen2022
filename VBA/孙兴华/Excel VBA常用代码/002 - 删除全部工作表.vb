Sub DelShet() '删除所有工作表
    Dim sht As Worksheet
    Application.ScreenUpdating = False '关屏幕刷新
    Application.DisplayAlerts = False '关警告信息
    On Error Resume Next
    For Each sht In Worksheets
        sht.Delete '遍历工作表删除
    Next
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub