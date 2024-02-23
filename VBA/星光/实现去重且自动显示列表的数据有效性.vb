Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Intersect([a:a], Target) Is Nothing Then Exit Sub
    '如果选择的单元格不存在于A列，则退出。A列是设置数据验证的区域
    If Target.Rows.Count > 1 Then Exit Sub
    '不允许选择多行
    Dim arr, brr, i&, j&, k&, s
    Dim d As Object
    Set d = CreateObject("scripting.dictionary")
    '后期绑定字典
    arr = Range("d1:d" & Cells(Rows.Count, "d").End(xlUp).Row)
    '数据来源列
    If Not IsArray(arr) Then Exit Sub
    '如果不存在数据源选项，则arr非数组，那么退出程序
    For i = 2 To UBound(arr)
        'D1是标题，从第2行开始遍历数据源，将人名装入字典
        If arr(i, 1) <> "" Then d(arr(i, 1)) = ""
    Next
    s = Join(d.keys, ",")
    With Target.Validation
         .Delete '删掉旧的
         .Add Type :=xlValidateList, AlertStyle:=xlValidAlertStop,  _
                Operator:=xlBetween, Formula1:=s
        's为数据验证的序列来源
    End With
    Application.SendKeys "%{down}"
    'SendKeys发出快捷键atl+↓直接弹出数据验证下拉列表
    Set d = Nothing
    '释放字典内存
End Sub