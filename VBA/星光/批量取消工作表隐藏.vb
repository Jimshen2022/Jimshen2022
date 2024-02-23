取消全部
Sub unShtVisible()
    Dim sht As Worksheet
    For Each sht In Worksheets '遍历工作表，设置可见
        sht.Visible = xlSheetVisible
    Next
End Sub


取消部分
Sub unShtVisible()
    Dim sht As Worksheet, t
    t = "看见星光/Excel星球/Sheet5/" '将需要隐藏的工作表名称写在这
    For Each sht In Worksheets '遍历工作表，设置可见
        If InStr(t, sht.Name & "/") Then
            sht.Visible = xlSheetVisible
        End If
    Next
End Sub