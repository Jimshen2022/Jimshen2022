
Sub 打印()
    qs = [I1]: zs = [I2] '起始、张数
    If zs Mod 2 = 0 Then zs1 = zs Else zs1 = zs - 1
    For i = 1 To zs1 Step 2 '打印到偶数张
        [e4] = qs + i - 1:[e19] = qs + i
        ActiveSheet.PrintPreview
        'ActiveSheet.PrintOut
    Next
    If zs > zs1 Then '奇数张，最后打印半张
        [e4] = qs + zs - 1
        [a16:e24] = "" '清空下半张
        ActiveSheet.PrintPreview
        'ActiveSheet.PrintOut
        [a1:e9].Copy[a16:e24] '恢复下半张
    End If
End Sub