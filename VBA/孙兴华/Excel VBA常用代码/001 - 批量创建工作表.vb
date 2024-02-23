Sub NewSht()
    Dim shtActive As Worksheet, sht As Worksheet
    Dim i As Long, strShtName As String
    On Error Resume Next '当代码出错时继续运行
    Set shtActive = ActiveSheet
    For i = 2 To shtActive.Cells(Rows.Count, 1).End(xlUp).Row
    '单元格A1是标题，跳过，从第2行开始遍历工作表名称
        strShtName = shtActive.Cells(i, 1).Value
        '工作表名强制转换为字符串类型
        Set sht = Sheets(strShtName)
        '当工作簿不存在工作表Sheets(strShtName)时，这句代码会出错，然后……
        If Err Then
        '如果代码出错，说明不存在工作表Sheets(t)，则新建工作表
            Worksheets.Add , Sheets(Sheets.Count)
            '新建一个工作表，位置放在所有已存在工作表的后面
            ActiveSheet.Name = strShtName
            '新建的工作表必然是活动工作表，为之命名
            Err.Clear
            '清除错误状态
        End If
    Next
    shtActive.Activate
    '重新激活原工作表
End Sub