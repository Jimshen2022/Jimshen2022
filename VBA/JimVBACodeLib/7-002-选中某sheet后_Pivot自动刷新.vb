
' worksheet active
'automatical refresh pivottables, thse code in sheets not in modules.

Private Sub Worksheet_Activate()
    Dim Table As PivotCache
    For Each Table In ThisWorkbook.PivotCaches
        Table.Refresh
    Next Table
    Range("b1").Value = Format(Now(), "yyyy-mm-dd hh:mm:ss")
End Sub
