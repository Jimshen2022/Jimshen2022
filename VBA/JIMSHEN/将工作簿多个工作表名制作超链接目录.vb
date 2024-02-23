'将工作簿多个工作表名制作超链接目录

Sub ml()
    Dim sht As Worksheet, i&, shtname$
    Columns(1).ClearContents
    Cells(1, 1).Value = "List"
    i = 1
    For Each sht In Worksheets
        shtname = sht.Name
        If sht.Name <> ActiveSheet.Name Then
            i = i + 1
            ActiveSheet.Hyperlinks.Add anchor:=Cells(i, 1), Address:="", SubAddress:="'" & shtname & "'!a1", TextToDisplay:=shtname
        End If
    Next

End Sub