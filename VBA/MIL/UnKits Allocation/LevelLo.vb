Sub POSort()
Dim StockItemRange As Range
Dim StockQtyRange As Range
Dim Item As Variant
Dim Qty As Double, ItemTemp As Range, QtyTemp As Range
Dim x, y As Double
Dim sht As Worksheet

Set shtP = Sheets("PO_List")
LastRow = shtP.Cells(shtP.Rows.Count, "A").End(xlUp).Row
shtP.Range("E5:E" & LastRow).ClearContents
shtP.Range("I5:I" & LastRow).ClearContents

'Sort
    Sheets("PO_List").Select
    Sheets("PO_List").Range("A4", "F" & LastRow).Select
    ActiveWorkbook.Sheets("PO_List").Sort.SortFields.Clear
    ActiveWorkbook.Sheets("PO_List").Sort.SortFields.Add Key:=Range( _
        "F4:F" & LastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Sheets("PO_List").Sort.SortFields.Add Key:=Range( _
        "A4:A" & LastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    
    Sheets("PO_List").Range("E3").Select
    Selection.Copy
    Sheets("PO_List").Range("E5:E" & LastRow).Select

    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
    Sheets("PO_List").Range("I3").Select
    Selection.Copy
    Sheets("PO_List").Range("I5:I" & LastRow).Select

    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
    Calculate
   
    With ActiveWorkbook.Sheets("PO_List").Sort
        .SetRange Range("A4:F" & LastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub

Sub POLoad2nd()
Dim StockItemRange As Range
Dim StockQtyRange As Range
Dim Item As Variant
Dim Qty As Double, ItemTemp As Range, QtyTemp As Range
Dim x, y As Double
Dim sht As Worksheet

Set shtP = Sheets("PO_List")
LastRow = shtP.Cells(shtP.Rows.Count, "A").End(xlUp).Row

Set onh = Sheets("ControlFile")
LastRowo = onh.Cells(onh.Rows.Count, "A").End(xlUp).Row

Set StockItemRange = onh.Range("d3:d" & LastRowo)
Set StockQtyRange = onh.Range("s3:s" & LastRowo)

For i = 5 To LastRow
Item = shtP.Range("B" & i).Value

x = Application.SumIf(StockItemRange, Item, StockQtyRange)
y = Application.SumIf(shtP.Range("B4 : B" & i - 1), Item, shtP.Range("G4 : G" & i - 1))

shtP.Range("G" & i).Value = Application.Min(x - y, Range("D" & i))

Next i

End Sub

Sub ControlLoad()
Dim StockItemRange As Range
Dim StockQtyRange As Range
Dim Item As Variant
Dim Qty As Double, ItemTemp As Range, QtyTemp As Range
Dim x, y As Double
Dim sht As Worksheet

Set sht = Sheets("ControlFile")

LastRow = sht.Cells(sht.Rows.Count, "A").End(xlUp).Row
If LastRow = 2 Then
LastRow = 3
End If

'Shift

For y = 3 To LastRow

    sht.Range("R" & y) = Mid(sht.Range("M" & y), 2, 1)

Next y

'Sort
    'Sheets("ControlFile").Select
    'sht.Range("A2", "R" & LastRow).Select
    'ActiveWorkbook.Worksheets("ControlFile").Sort.SortFields.Clear
    'ActiveWorkbook.Worksheets("ControlFile").Sort.SortFields.Add Key:=Range( _
        "K3:K" & LastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    'ActiveWorkbook.Worksheets("ControlFile").Sort.SortFields.Add Key:=Range( _
        "R3:R" & LastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    'With ActiveWorkbook.Worksheets("ControlFile").Sort
        '.SetRange Range("A2:R" & LastRow)
        '.Header = xlYes
        '.MatchCase = False
        '.Orientation = xlTopToBottom
        '.SortMethod = xlPinYin
        '.Apply
    'End With
    

'Level Load

Set onh = Sheets("PO_List")

LastRowo = onh.Cells(onh.Rows.Count, "A").End(xlUp).Row

'Set pri = Sheets("0897")

'LastRowp = pri.Cells(pri.Rows.Count, "A").End(xlUp).Row

Set StockItemRange = onh.Range("B5:B" & LastRowo)
Set StockQtyRange = onh.Range("D5:D" & LastRowo)

For i = 3 To LastRow

Item = sht.Range("D" & i).Value

x = Application.SumIf(StockItemRange, Item, StockQtyRange)
y = Application.SumIf(sht.Range("D2 : D" & i - 1), Item, sht.Range("S2 : S" & i - 1))

sht.Range("S" & i).Value = Application.Min(x - y, Range("I" & i))

Next i

'Format

Sheets("PO_List").Select
Call POLoad2nd

Sheets("ControlFile").Select
MsgBox "LevelLoad Finished"
End Sub

