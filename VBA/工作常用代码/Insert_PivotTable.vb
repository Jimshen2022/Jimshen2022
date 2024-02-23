
Sub InsertOrderStatusPivot()
    'Macro By ExcelChamps 增加pivotTable到fulfillment Sheet
    
    '    Application.ScreenUpdating = False
    '    Application.Calculation = xlCalculationManual
    '    Application.StatusBar = "Calculating, please wait ......"
    
    
    'Declare Variables
    Dim PSheet As Worksheet
    Dim DSheet As Worksheet
    Dim PCache As PivotCache
    'Dim PTable As PivotTable  ？？不能理解，为何不要定义pivotTable
    Dim PRange As Range
    Dim LastRow As Long
    Dim LastCol As Long
    
    'Insert a New Blank Worksheet
    
    Set PSheet = Worksheets("Fulfillment")
    Set DSheet = Worksheets("Unpick_orders")
    Worksheets("Fulfillment").Activate
    On Error Resume Next
    ActiveSheet.PivotTables("PivotTable1").TableRange2.Clear
    
    'Define Data Range
    LastRow = DSheet.Cells(Rows.Count, 1).End(xlUp).Row
    LastCol = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    Set PRange = DSheet.Cells(1, 1).Resize(LastRow, LastCol)
    
    
    'Define Pivot Cache '确定pivottable的位置 cells(3,1)
    Set PCache = ActiveWorkbook.PivotCaches.Create _
            (SourceType:=xlDatabase, SourceData:=PRange). _
            CreatePivotTable(TableDestination:=PSheet.Cells(3, 1),  _
            TableName:="PivotTable1")
    
    'Insert Blank Pivot Table
    Set PTable = PCache.CreatePivotTable _
            (TableDestination:=PSheet.Cells(1, 1), TableName:="PivotTable1")
    
    'Insert Row Fields    '增加row值
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Order Avaiable STO Fulfillment")
         .Orientation = xlRowField
         .Position = 1
    End With
    
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("No")
         .Orientation = xlRowField
         .Position = 2
    End With
    
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Intervals")
         .Orientation = xlRowField
         .Position = 3
    End With
    
    'Insert Column Fields    '增加column值
    'With ActiveSheet.PivotTables("PivotTable1").PivotFields("")
    '.Orientation = xlColumnField
    '.Position = 1
    'End With
    
    'Insert Data Field
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("RP Order#")
         .Orientation = xlDataField
         .Position = 1
         .Function  = xlSum
             .NumberFormat = "#,##0"
        End With
        
        With ActiveSheet.PivotTables("PivotTable1").PivotFields("QTY")
             .Orientation = xlDataField
             .Position = 2
             .Function  = xlSum
                 .NumberFormat = "#,##0"
            End With
            
            With ActiveSheet.PivotTables("PivotTable1").PivotFields("QTY")
                 .Orientation = xlDataField
                 .Position = 3
                 .Calculation = xlPercentOfTotal
                 .NumberFormat = "0.00%"
                 .Caption = "%(QTY)"
            End With
            
            'Format Pivot Table
            ActiveSheet.PivotTables("PivotTable1").ShowTableStyleRowStripes = True
            ActiveSheet.PivotTables("PivotTable1").TableStyle2 = "PivotStyleMedium9"
            ActiveSheet.PivotTables("PivotTable1").MergeLabels = True '
            
            ''经典格式
            With ActiveSheet.PivotTables("PivotTable1")
                 .InGridDropZones = True
                 .RowAxisLayout xlTabularRow
            End With
            
            With ActiveSheet.PivotTables("PivotTable1")
                 .PivotFields("NO").Subtotals(1) = False
                 .PivotFields("Intervals").Subtotals(1) = False
            End With
            
            '置中对齐
            Columns("D:F").Select
            With Selection
                 .HorizontalAlignment = xlCenter
                 .Orientation = 0
                 .AddIndent = False
                 .IndentLevel = 0
                 .ShrinkToFit = False
                 .ReadingOrder = xlContext
                 .MergeCells = False
            End With
            
            'don't auto fit column width on updated
            Columns("F:F").ColumnWidth = 13.71
            ActiveSheet.PivotTables("PivotTable1").HasAutoFormat = False
            
            
            '    Application.Calculation = xlCalculationAutomatic
            '    Application.ScreenUpdating = True
            '    Application.StatusBar = False
            
            
        End Sub
