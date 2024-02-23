Sub refresh_piovtTables()
    Dim Table As PivotCache
    For Each Table In ThisWorkbook.PivotCaches
        Table.Refresh
    Next Table
    
End Sub


Sub Pivot_Refresh3()
    
    Dim Table As PivotTable
    Set Table = ActiveSheet.PivotTables("Customer Data")
    
    Table.RefreshTable
    
End Sub
