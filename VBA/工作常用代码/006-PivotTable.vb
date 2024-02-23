

'D:\Document\06-Millennium\00-Report\03-PickingException\MIL Picking exception - 2022
Sub PivotTabel_DATA()
    
    Dim wksData As Worksheet
    Dim objCache As PivotCache
    Dim objTabel As PivotTable
    Dim Arr() As Variant
    
    With Sheet1
        Set wksData = Sheet1
        Arr = wksData.Range("a1").CurrentRegion
        Sheet3.Cells.ClearContents
        Set objCache = ThisWorkbook.PivotCaches.Create(xlDatabase, wksData.Range("a1").CurrentRegion.Address(External:=True))
        Set objTabel = objCache.CreatePivotTable(Sheet3.Range("a3"))
        With objTabel
             .AddFields RowFields:=Array(Arr(1, 25), Arr(1, 24)), _
                    ColumnFields:=Array(Arr(1, 16))
                    'PageFields:=Array(Arr(1, 3))
             .AddDataField .PivotFields(Arr(1, 11)), , xlSum '
             .RowAxisLayout xlOutlineRow
             .TableStyle2 = "PivotStyleDark7"    'desgin format
            .PivotFields("Sum of Transaction Quantity").NumberFormat = "#,##0"    'VALUE FORMAT
            .PivotFields("Reason").AutoSort xlDescending, "Sum of Transaction Quantity"    'SORTING
            .InGridDropZones = True      'display as classic format type
            .RowAxisLayout xlTabularRow
            .MergeLabels = True    'merge the row data
        End With
        Sheet3.Activate
        Range("A:A").Columns.ColumnWidth = 26
        Range("B:B").Columns.ColumnWidth = 50
        Range("C:M").Columns.ColumnWidth = 9.57
        Range("C:M").Columns.HorizontalAlignment = xlCenter
        Range("C:M").Columns.VerticalAlignment = xlBottom
        
    End With
End Sub



'v02 pivotTable


Sub NewPivoTable66()

    Dim wksData As Worksheet
    Dim objCache As PivotCache
    Dim objTable As PivotTable
    Dim avntArr() As Variant
    Sheet3.Cells.ClearContents
    Set wksData = Worksheets("DATA")
    avntArr = wksData.Range("A1:U1")
    Set objCache = ThisWorkbook.PivotCaches.Create( _
        xlDatabase, wksData.Range("a1").CurrentRegion. _
        Address(External:=True))
    Set objTable = objCache.CreatePivotTable _
        (Sheet3.Range("a3"))
    With objTable
        .AddFields RowFields:=Array(avntArr(1, 8), avntArr(1, 7)), _
            ColumnFields:=Array(avntArr(1, 6)), _
            PageFields:=Array(avntArr(1, 3), avntArr(1, 1))
        .AddDataField .PivotFields(avntArr(1, 9)), , xlCount

        .RowAxisLayout xlOutlineRow
        .TableStyle2 = "PivotStyleDark7"    'desgin format
        
      'Kind filter "UPH"
      .PivotFields("Kind").CurrentPage = "(All)"
      .PivotFields("Kind").PivotItems("Others").Visible = False
      .PivotFields("Kind").EnableMultiplePageItems = True
        
			
        .PivotFields("Count of Serial Number").NumberFormat = "#,##0"    'VALUE FORMAT
        .PivotFields("LocationType").AutoSort xlDescending, "Count of Serial Number"    'SORTING
        
    End With
    
    
    'Format
    With Sheet3
        Sheet3.Activate
        .Range("d1") = "DataCollectedAt:  " & Format(Now, "HH:MM:SSam/pm,mmm.dd.yyyy")
        .Columns("C:F").Select
        Selection.ColumnWidth = 16
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With

        
    End With
    
'    With objTable.PivotFields(avntArr(1, 9))
'         .NumberFormat = "#,##0"
'    End With
            
    Set wksData = Nothing
    Set objCache = Nothing
    Set objTable = Nothing
End Sub