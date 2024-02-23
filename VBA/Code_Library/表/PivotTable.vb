Sub NewPivotTabel_Ashley_SH() 'Asia_OnHand -20210221.1309.xlsx
    
    Dim wksData As Worksheet
    Dim objCache As PivotCache
    Dim objTabel As PivotTable
    Dim Arr() As Variant
    
    With Sheet11
        Set wksData = Sheet11
        Arr = wksData.Range("a1").CurrentRegion
        Set objCache = ThisWorkbook.PivotCaches.Create(xlDatabase, wksData.Range("a1").CurrentRegion.Address(external:=True))
        Set objTabel = objCache.CreatePivotTable(wksData.Range("m3"))
        With objTabel
             .AddFields RowFields:=Array(Arr(1, 10)),  _
                    ColumnFields:=Array(Arr(1, 2)),  _
                    PageFields:=Array(Arr(1, 3))
             .AddDataField .PivotFields(Arr(1, 4)), , xlSum '
             .RowAxisLayout xlOutlineRow
        End With
    End With
    
End Sub