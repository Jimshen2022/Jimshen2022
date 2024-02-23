Sub unfilter2()
On Error Resume Next

Application.ScreenUpdating = False
Dim i%, sht As Worksheet

For Each sht In Worksheets
    If sht.AutoFilterMode = True Then sht.Range("a1:zz1").AutoFilter
    If sht.AutoFilterMode = False Then sht.Range("a1:zz1").AutoFilter Field:=1
Next

Application.ScreenUpdating = True
End Sub




