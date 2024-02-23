

Private Sub Worksheet_Activate()
    ActiveSheet.PivotTables("PivotTable1").PivotCache.Refresh
    ActiveSheet.PivotTables("PivotTable9").PivotCache.Refresh
End Sub




Sub refresh_piovtTables()
    Dim Table As PivotCache
    For Each Table In ThisWorkbook.PivotCaches
        Table.Refresh
    Next Table
    
End Sub
