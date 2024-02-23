Sub delsht()
    Rem 删除非活动工作表
    Dim sht As Worksheet
    Application.DisplayAlerts = False
    For Each sht In Worksheets
        If sht.Name <> ActiveSheet.Name Then
            sht.Delete
        End If
    Next
    Application.DisplayAlerts = True
End Sub