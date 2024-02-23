'范例39 新建工作簿

Sub AddNowwb()
    Dim AddNowwb As Workbook
    Dim shtname As Variant
    Dim arr As Variant
    Dim i As Integer
    Dim MyInNewWb As Integer
    
    MyInNewWb = Application.SheetsInNewWorkbook
    arr = Array("品名", "单价", "数量", "金额")
    shtname = Array("01月", "02月", "03月", "04月", "05月", "06月", "07月", "08月", "09月", "10月", "11月", "12月")
    Application.SheetsInNewWorkbook = 12
    Set AddNowwb = Workbooks.Add
    With AddNowwb
        For i = 1 To 12
            With .Sheets(i)
                .Name = shtname(i - 1)
                .Range("a1").Resize(1, UBound(arr) + 1) = arr
            End With
        Next
        .SaveAs FileName:=ThisWorkbook.Path & "\" & "存货明细.xlsx"
        .Close savechanges:=True
    End With
    Application.SheetsInNewWorkbook = MyInNewWb
    Set AddNowwb = Nothing
    
End Sub