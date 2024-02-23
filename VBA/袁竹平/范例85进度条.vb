' 第6章 控件与用户窗体


sub jdt()
rs.Open sql,  connXls, 1
    Dim p As Integer: p = 0
    Do While Not rs.EOF
        p = p + 1
        '在状态栏显示
        Application.StatusBar = GetProgress(p, rs.RecordCount)
    ……

end sub


'自定义的进度条，在状态栏显示
Function GetProgress(curValue, maxValue)
Dim i As Single, j As Integer, s As String
i = maxValue / 20
j = curValue / i
 
For m = 1 To j
    s = s & "■"
Next m
For n = 1 To 20 - j
    s = s & "□"
Next n
GetProgress = s & FormatNumber(curValue / maxValue * 100, 2) & "%"
End Function



sub jdt2()
Dim x               As Integer 
Dim MyTimer         As Double 
 
'Change this loop as needed.
For x = 1 To 50
    ' Do stuff
    Application.StatusBar = "Progress: " & x & " of 50: " & Format(x / 50, "0%")
Next x 
 
Application.StatusBar = False
end Sub