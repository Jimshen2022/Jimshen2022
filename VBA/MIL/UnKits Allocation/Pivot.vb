Sub WNPivotTable()

    Dim PTable As PivotTable
    Dim PCache As PivotCache
    Dim PRange As Range
    Dim PSheet As Worksheet
    Dim DSheet As Worksheet
    Dim LR As Long
    Dim LC As Long
    
    On Error Resume Next
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Worksheets("WNPivot").Delete 'This will delete the exisiting pivot table worksheet
    Worksheets.Add After:=ActiveSheet ' This will add new worksheet
    ActiveSheet.Name = "WNPivot" ' This will rename the worksheet as "Pivot Sheet"
    On Error GoTo 0
    
    Set PSheet = Worksheets("WNPivot")
    Set DSheet = Worksheets("LINK")
    
    'Find Last used row and column in data sheet
    LR = DSheet.Cells(Rows.Count, 1).End(xlUp).Row
    LC = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    
    'Set the pivot table data range
    Set PRange = DSheet.Cells(1, 1).Resize(LR, LC)
    
    'Set pivot cache
    Set PCache = ActiveWorkbook.PivotCaches.Create(xlDatabase, SourceData:=PRange)
      
   'Create blank pivot table
    Set PTable = PCache.CreatePivotTable(TableDestination:=PSheet.Cells(1, 1), TableName:="MIL")
    
    'With ActiveSheet.PivotTables("WN").PivotFields("Note")
        '.Orientation = xlPageField
        '.Position = 1
    'End With
    
    'ActiveSheet.PivotTables("WN").PivotFields("Note").ClearAllFilters
    'ActiveSheet.PivotTables("WN").PivotFields("Note").CurrentPage = "Short"
    
    'Insert country to Row Filed
    
    With PSheet.PivotTables("MIL").PivotFields("Customer")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    With PSheet.PivotTables("MIL").PivotFields("Shift")
        .Orientation = xlRowField
        .Position = 2
    End With
    
     
    With PSheet.PivotTables("MIL").PivotFields("DUEDATE")
        .Orientation = xlColumnField
        .Position = 1
    End With
      
    'Insert Sales column to the data field
    
    With PSheet.PivotTables("MIL").PivotFields("Box")
        .Orientation = xlDataField
        .Position = 1
    End With
              
    ActiveSheet.PivotTables("MIL").ShowDrillIndicators = False
    
    'Format Pivot Table
    PSheet.PivotTables("MIL").ShowTableStyleRowStripes = True
    PSheet.PivotTables("MIL").TableStyle2 = "PivotStyleMedium14"
    
    'Show in Tabular form
    PSheet.PivotTables("MIL").RowAxisLayout xlTabularRow
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    

    
    With ActiveSheet.PivotTables("MIL").PivotFields("Sum of Box")
        .Caption = "TotalBox"
    End With
    
    'Call NoSubtotalsWN
    endrow = PSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    Sheets("WNPivot").Select
    With ActiveWorkbook.Sheets("WNPivot").Tab
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.399975585192419
    End With
    
    Call Format
End Sub

Sub NoSubtotalsWN()

Dim pt As PivotTable
Dim pf As PivotField
On Error Resume Next
For Each pt In ActiveSheet.PivotTables
  pt.ManualUpdate = True
  
  For Each pf In pt.PivotFields
    'First, set index 1 (Automatic) to True,
    'so all other values are set to False
    pf.Subtotals(1) = True
    pf.Subtotals(1) = False
  Next pf
  pt.ManualUpdate = False

Next pt

ActiveSheet.PivotTables("MIL").ColumnGrand = False

Columns("A:D").Select
Columns("A:D").EntireColumn.AutoFit
    
End Sub


