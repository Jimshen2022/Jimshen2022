Sub GetShName()
    Dim sht As Worksheet, k As Long
    Application.ScreenUpdating = False
    With Range("a:a")
         .Clear '清除所有
         .NumberFormat = "@" '设置文本格式
    End With
    k = 1
    Cells(1, 1) = "目录"
    For Each sht In Sheets '遍历工作表
        k = k + 1 '累加个数
        Cells(k, 1) = sht.Name
    Next
    Application.ScreenUpdating = True
End Sub






Sub NewShName()
    Dim aData, aRes, i As Long
    If ActiveWorkbook.ProtectStructure = True Then
        MsgBox "工作簿有保护，工作表无法重命名。"
        Exit Sub
    End If
    Application.ScreenUpdating = False
    On Error Resume Next '忽略错误，继续运行
    aData = Range("a1:b" & Cells(Rows.Count, 1).End(xlUp).Row)
    ReDim aRes(1 To UBound(aData), 1 To 1)
    For i = 1 To UBound(aData)
        Err.Clear '错误状态清除
        If aData(i, 2) <> "" Then
            Sheets(aData(i, 1)).Name = aData(i, 2)
            If Err.Number Then '如果有错
                aRes(i, 1) = "更名失败"
            Else
                aRes(i, 1) = "成功"
            End If
        Else
            aRes(i, 1) = "空白值"
        End If
    Next
    Range("c1").Resize(UBound(aRes), 1) = aRes
    Application.ScreenUpdating = True
    MsgBox "更名完成，结果参考C列。"
End Sub