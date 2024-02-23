
‘1210511 AS400这类字符串可以用如下公式转换成日期格式
brr(i, 20) = CDate(Format("20" & Mid(brr(i, 6), 2, 2) & "/" & Mid(brr(i, 6), 4, 2) & "/" & Mid(brr(i, 6), 6, 2), "mm/dd/yyyy"))



Sub zldccmx()
    Dim a As Date, b As Date
    a = CDate(Format(Now - 1, "yyyy-mm-dd") & " " & "18:30:00") '昨天18：30：00
    b = CDate(Format(Now, "yyyy-mm-dd") & " " & "09:00:00") '今天9：00：00
    Debug.Print "相隔" & Int(24 * (b - a)) & "小时" & Format(b - a, "nn分ss秒")
End Sub




Sub AccumulateUtilization()
    
    Dim i&, nrow%
    Application.ScreenUpdating = False
    With Sheet4
         .Range("i1") = "Date"
        
        nrow =  .Range("a1048576").End(3).Row
        
        For i = 2 To nrow
            If  .Cells(i, 1) = "" Then
                GoTo 100
            ElseIf  .Cells(i, 1).Value <> "" Then
                 .Cells(i, 9) = CDate(Format(Now(), "yyyy/mm/dd"))
            End If
100
        Next
    End With
    
    Application.ScreenUpdating = True
End Sub
