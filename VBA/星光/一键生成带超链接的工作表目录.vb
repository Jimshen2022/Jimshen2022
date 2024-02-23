Sub SheetList()
    Dim sht As Worksheet, i As Long, strName As String
    With Columns(1)
         .Clear '清空A列数据
         .NumberFormat = "@" '设置文本格式
    End With
    Range("a1") = "目录"
    For i = 1 To Sheets.Count '索引法遍历工作表集合
        strName = Sheets(i).Name '表名
        Cells(i + 1, 1).Value = strName
        ActiveSheet.Hyperlinks.Add anchor:=Cells(i + 1, 1), Address:="",  _
                SubAddress:="'" & strName & "'!a1", TextToDisplay:=strName
    Next
End Sub