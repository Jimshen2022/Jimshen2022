Sub LPQty() 'SHEET7 COLUMN AE and AD
    
    '    Application.ScreenUpdating = False
    '    Application.Calculation = xlCalculationManual
    '    Application.StatusBar = "Counting orders, please wait ......"
    
    Sheet7.Activate
    'Range("s2:s66365").ClearContents
    Dim i%, row1%
    For i = 2 To [a66365].End(3).Row
        row1 = Application.WorksheetFunction.CountIfs(Range("ac2:ac" & i), Cells(i, "ac"))
        If row1 = 1 Then Cells(i, "ae") = 1
        If row1 > 1 Then Cells(i, "ae") = 0 'countifs 如果AC栏有重复为0，否则为1
        Cells(i, "ad") = Application.WorksheetFunction.WeekNum(Cells(i, "t")) 'weeknum
    Next i
    
    '    Application.Calculation = xlCalculationAutomatic
    '    Application.ScreenUpdating = True
    '    Application.StatusBar = False
    
End Sub