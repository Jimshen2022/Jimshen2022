Sub PivotTabel_MIL_OH()
    
    Application.ScreenUpdating = False
    Dim wksData As Worksheet
    Dim objCache As PivotCache
    Dim objTabel As PivotTable
    Dim arr() As Variant
    
    Sheet5.Cells.Clear
    
    With Sheet1
        Set wksData = Sheet1
        arr = wksData.Range("a1").CurrentRegion
        Set objCache = ThisWorkbook.PivotCaches.Create(xlDatabase, wksData.Range("a1").CurrentRegion.Address(external:=True))
        Set objTabel = objCache.CreatePivotTable(Sheet5.Range("a5"))
        
'        With objTabel
'             .AddFields RowFields:=Array(Arr(1, 10)), _
'                    ColumnFields:=Array(Arr(1, 2)), _
'                    PageFields:=Array(Arr(1, 3))
'             .AddDataField .PivotFields(Arr(1, 4)), , xlSum '
'             .RowAxisLayout xlOutlineRow
'        End With
        
        With objTabel

             .AddFields ColumnFields:=Array(arr(1, 10)), _
                    PageFields:=Array(arr(1, 3))
             .AddDataField .PivotFields(arr(1, 7)), , xlSum
             .AddDataField .PivotFields(arr(1, 12)), , xlSum
             .DataPivotField.Orientation = xlRowField
             .DataPivotField.Position = 1
             .RowAxisLayout xlOutlineRow
             .PivotFields("Sum of LQNTY").NumberFormat = "#,##0"
             .PivotFields("Sum of AMT($USD)").NumberFormat = "#,##0"
             .PivotFields("Product").PivotItems("UPH").Position = 1
             .PivotFields("Product").PivotItems("CG").Position = 2
             .PivotFields("Product").PivotItems("Bedding").Position = 3
             .PivotFields("Product").PivotItems("ZipperCover").Position = 4
             .PivotFields("Product").PivotItems("UnKits").Position = 5
             .PivotFields("Product").PivotItems("RP").Position = 6
             .PivotFields("Product").PivotItems("Plastics").Position = 7
             .PivotFields("Product").PivotItems("Foundation").Position = 8
             .PivotFields("Product").PivotItems("RawMaterial").Position = 9
             .PivotFields("Product").PivotItems("Verona").Position = 10
             .PivotFields("Product").PivotItems("Panel").Position = 11
        
        End With
        
        
    With Sheet5.Columns("b:m")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .Range("a1").Value = "DataCollectedAt:  " & Format(Now, "HH:MM:SSam/pm,mmm.dd.yyyy")
        .Range("a1").Font.Color = -16776961
    End With
        
    End With
    Application.ScreenUpdating = True
End Sub

