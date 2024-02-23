Sub AddBarCode()
    Application.ScreenUpdating = False
    On Error Resume Next
    Dim rng As Range, rngData As Range, sht As Worksheet
    Set rngData = Application.InputBox( _
                    "请选择需要生成条形码的区域。", _
                    "Excel VBA", _
                    Type:=8)
    Set sht = rngData.Parent '所选区域的工作表
    Set rngData = Intersect(rngData, sht.UsedRange) '交集运算取实际区域
    If rngData Is Nothing Or Err Then
        MsgBox "你未选择有效的区域，程序退出。"
        Exit Sub
    End If
    sht.Select
    For Each shp In sht.Shapes '删除表格原有条形码
        If InStr(shp.Name, "BarCodeCtrl") Then shp.Delete
    Next
    For Each rng In rngData
        With sht.OLEObjects.Add(classtype:="BARCODE.BarCodeCtrl.1")
             .Object.Style = 6 '二维码修改为11
             '.Object.Value = rng.Value
             .LinkedCell = rng.Address '链接单元格
             .Height = rng.Height - 2
             .Width = rng.Width - 2
             .Left = rng.Left
             .Top = rng.Top
         End With
    Next
    Application.ScreenUpdating = True
End Sub