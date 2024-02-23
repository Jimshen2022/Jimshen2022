'工作表加密后却把密码忘记了 ？没事 ，秒破 ！

Sub UnProtct()
    Dim sht As Worksheet, shtAct As Worksheet
    MsgBox "破解提示：当要求输入密码时请点击取消！"
    Application.DisplayAlerts = False
    On Error Resume Next
    Set shtAct = ActiveSheet
    For Each sht In Worksheets
        With sht
             .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True,  _
                    AllowFiltering:=True, AllowUsingPivotTables:=True
             .Protect DrawingObjects:=False, Contents:=True, Scenarios:=False,  _
                    AllowFiltering:=True, AllowUsingPivotTables:=True
             .Protect DrawingObjects:=True, Contents:=True, Scenarios:=False,  _
                    AllowFiltering:=True, AllowUsingPivotTables:=True
             .Protect DrawingObjects:=False, Contents:=True, Scenarios:=True,  _
                    AllowFiltering:=True, AllowUsingPivotTables:=True
             .Unprotect
        End With
    Next
    shtAct.Select
    Application.DisplayAlerts = True
    MsgBox "ok"
End Sub