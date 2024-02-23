Sub AshleySHRPOpenOrdersFulfillment()

    t = Timer
    'ActiveWorkbook.SaveAs ActiveWorkbook.Path & "\Ashley SH RP Open Orders Fulfillment-" & Format(Now(), "yyyymmdd.hhmm") & ".xlsb"
    Application.ScreenUpdating = False
    Sheet6.Range("i1").Value = "Data collected at:" & Format(Now(), "hhmm,mm-dd-yyyy")

    Call unfilter2
    Call RpOpenOrderLoading
    Call OrderStatusJudge
    Call RPOrderCounted
    Call rp_creation_date_pending_days
    Call Column_V_Intervals
    Call intervals_no
    
    Call load_mapics_on_hand
    Call picked_and_packed
    
    Call loading_unpick_orders
    Call ColumnX
    Call ColumnYZAA
    Call Order_fulfillment
    Call LoadPO
    Call InsertOrderStatusPivot
    Call InsertSummaryPivot1
    Call InsertSummaryPivot2
    Call HaveStockRPorderList
    Call Shortage_vs_PO
    Call POAllocationForShortage
    Call PullASyardSTO
    Sheet6.Range("a1").Value = "Data collected at:  " & Format(Now(), "hh:mm am/pm,mmm.dd.yyyy")
    Sheet6.Select
    Application.ScreenUpdating = True
    MsgBox "Updated Successful~    " & Format(Timer - t, "0.00" & "s")
End Sub



Sub POAllocationForShortage()   'ready 12  allocation PO and MO for Trips

'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
'    Application.StatusBar = "Calculate picked and packed RP Orders, please wait ......"

Dim arr1(), arr2()
Dim a%, b%

Sheet10.Activate   'shortageSheet
'Range("b1", Cells(Rows.Count, "g").End(xlUp)).Sort [d1], xlAscending, [b1], , xlAscending, , , xlYes  '按款号与日期排序
arr1 = Range("a2", Cells(Cells(Rows.Count, "n").End(xlUp).Row, "o"))    '数组赋值
Sheet2.Activate   ' supplySheet
'Range("a1", Cells(Rows.Count, "c").End(xlUp)).Sort [a1], xlAscending, [c1], , xlAscending, , , xlYes   '按款号与日期排序
arr2 = Range("a2", Cells(Rows.Count, "y").End(xlUp)) '数组赋值

'For a = LBound(arr1) To UBound(arr1)
'    arr1(a, 6) = arr1(a, 6) * 1   '将文字改为数值
'Next

Sheet10.Activate
Range("o2:o66365").ClearContents

100
            For a = LBound(arr1) To UBound(arr1)
            For b = LBound(arr2) To UBound(arr2)
                If arr1(a, 5) = arr2(b, 2) And arr1(a, 14) <= arr2(b, 3) And arr1(a, 14) <> 0 Then    '订单款号与PO款号相同，且 订单需求<=PO数量 且 订单需求不为0
                    arr2(b, 3) = arr2(b, 3) - arr1(a, 14)                'PO数量 = PO数量 - 订单需求
                    arr1(a, 14) = arr1(a, 14) - arr1(a, 14)                '订单需求  = 订单需求 - 订单需求
                    Sheet10.Cells(a + 1, "o") = arr2(b, 1)            '预计齐料日 = PO的预计进料日期   ‘cells(a+1)是为了跟数组第一行对齐
                    arr1(a, 15) = arr2(b, 1)
                ElseIf arr1(a, 5) = arr2(b, 2) And arr1(a, 14) > arr2(b, 3) And arr2(b, 3) <> 0 Then  '亦或，订单款号与PO款号相同，且 订单需求>PO数量 且 PO不为0
                    arr1(a, 14) = arr1(a, 14) - arr2(b, 3)       '订单缺货数量 = 订单缺货数量 - PO数量
                    arr2(b, 3) = arr2(b, 3) - arr2(b, 3)       'PO数量 = PO数量 - PO数量
                    If arr1(a, 14) <> 0 Then                    '如果订单需求<>0, 回到100
                    GoTo 100
                    End If
                End If
            Next
            Next
        


Sheet10.Select

Dim i%
    For i = 2 To [M66365].End(3).Row
        If Cells(i, "O") = "" Then Cells(i, "O") = "PO Uncovered"
    Next

    Columns("N:O").Select
    Range("O1").Activate
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
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
    Range("O11").Select


'    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True
'    Application.StatusBar = False

End Sub



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
CreatePivotTable(TableDestination:=PSheet.Cells(3, 1), _
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
.Function = xlSum
.NumberFormat = "#,##0"
End With

With ActiveSheet.PivotTables("PivotTable1").PivotFields("QTY")
.Orientation = xlDataField
.Position = 2
.Function = xlSum
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
ActiveSheet.PivotTables("PivotTable1").MergeLabels = True  '

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


Sub HaveStockRPorderList()

'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
'    Application.StatusBar = "Calculating, please wait ......"


Sheet7.Activate
Range("a1:k66653").ClearContents

Set CNN = CreateObject("adodb.connection")
CNN.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=Excel 12.0;Data Source=" & ThisWorkbook.FullName
Sql = "select [RPKEY],[CUSNO],[CUSNM],[ITEMNO],[ITDSC],[QTY],[MODEL],[Type],[Order Creation Date],[Pending days],[Order Avaiable STO Fulfillment] " _
    & "from [Unpick_orders$] " _
    & "where [Order Avaiable STO Fulfillment]=""HaveStock"" " _
    & "Order by [RPKEY],[ITEMNO]"

Sheet7.Range("a2").CopyFromRecordset CNN.Execute(Sql)
CNN.Close
Set CNN = Nothing

Range("a1:k1") = Array("RPKEY", "CUSNO", "CUSNM", "ITEMNO", "ITDSC", "QTY", "Model", "Type", "Order Creation Date", "Pending days", "Order Avaiable STO Fulfillment")
Columns("i:i").NumberFormat = "m/d/yyyy"
Columns("A:k").EntireColumn.AutoFit

Range("A1:K1").Select
    Range("K1").Activate
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
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
    Columns("A:K").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
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
'
'    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True
'    Application.StatusBar = False



End Sub


Sub load_mapics_on_hand()

'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
'    Application.StatusBar = "Loading Mapics On Hand, please wait ......"

    Dim i As Long
    Dim adors As New Recordset
    Sheets("mapics_onhand").Activate
    Sheets("mapics_onhand").Range("a2:g66365").ClearContents
    
    
    Set Db = New Connection
    Db.CursorLocation = adUseClient
    If Db.State = 1 Then Db.Close
    
    UName = Sheet11.Range("a1")
    UPass = Sheet11.Range("a2")
    DateStart = Sheet11.Range("a3")
    DateEnd = Sheet11.Range("a4")
    
   
    Db.Open "Provider =IBMDASQL.DataSource.1" & _
     ";Catalog Library List=JIMTDTA" & _
     ";Persist Security Info=True" & _
     ";Force Translate=0" & _
     ";Data Source = WVFHA" & _
     ";User ID = " & UName & "" & _
     ";Password = " & UPass
     
     Set adors = New Recordset
     If adors.State = 1 Then adors.Close

    cmdtxt = "SELECT ITEMBL.ITNBR, ITEMBL.HOUSE, ITEMBL.ITCLS, ITEMBL.MOHTQ, ITEMBL.WHSLC, ITEMBL.QTSYR, ITMRVA.ITDSC " & _
             "FROM AMFLIBQ.ITEMBL ITEMBL, AMFLIBQ.ITMRVA ITMRVA, AMFLIBQ.WHSMST WHSMST " & _
             "WHERE ITMRVA.ITCLS = ITEMBL.ITCLS AND ITMRVA.ITNBR = ITEMBL.ITNBR AND ITMRVA.STID = WHSMST.STID AND ITEMBL.HOUSE = WHSMST.WHID AND ((ITEMBL.HOUSE='PC1' OR ITEMBL.HOUSE='232') AND (ITEMBL.MOHTQ<>0)) " & _
             "ORDER BY ITEMBL.ITNBR "

    adors.Open cmdtxt, Db, 3, 3
     For i = 0 To adors.Fields.Count - 1
         Worksheets("mapics_onhand").Cells(1, i + 1) = adors.Fields(i).Name
     Next i
     
     Worksheets("mapics_onhand").Range("a2").CopyFromRecordset adors
     adors.Close
     Set adors = Nothing
    
    
    Sheet4.Columns("A:D").NumberFormat = "@"
    
'    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True
'    Application.StatusBar = False
    
    
    

End Sub


Sub picked_and_packed()    '字典数组法

'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
'    Application.StatusBar = "Calculate picked and packed RP Orders, please wait ......"

    Dim d As Object
    Dim ar As Variant, br As Variant
    Set d = CreateObject("scripting.dictionary")
    r = Sheet4.Cells(Rows.Count, 1).End(3).Row
    Sheet4.Range("h2:i66653").ClearContents
       
    ar = Sheet4.[a1].CurrentRegion   '给数组ar赋值
    For i = 2 To UBound(ar)
        If Trim(ar(i, 1)) <> "" Then
            d(Trim(ar(i, 1))) = i
        End If
    Next i
    br = Sheet1.[a1].CurrentRegion    '给数组br赋值
    For i = 2 To UBound(br)
        If Trim(br(i, 18)) = "Packed and waiting for stick on trip" Or Trim(br(i, 18)) = "Stuck on trip and waiting for picking" Then
            If Trim(br(i, 11)) <> "" Then
                m = d(Trim(br(i, 11)))  '给字典赋值
                If m <> "" Then ar(m, 8) = ar(m, 8) + br(i, 13)
                End If
            End If
    Next i
    Sheet4.[a1].CurrentRegion = ar
    
    Dim x%
        For x = 2 To [a66563].End(3).Row
            If Cells(x, "h") = "" Then Cells(x, "h") = 0
            Next x
            
            
    Dim a%
        For a = 2 To [a66563].End(3).Row
            If Cells(a, "d") - Cells(a, "h") < 0 Then Cells(a, "i") = 0 Else Cells(a, "i") = Cells(a, "d") - Cells(a, "h")
            If Cells(a + 1, "a") = Cells(a, "a") Then Cells(a + 1, "i") = 0
        Next a
                
    Dim y%
        For y = 2 To [a66563].End(3).Row
            
            If Cells(y + 1, "a") = Cells(y, "a") Then Cells(y + 1, "i") = Cells(y + 1, "d")
            
                        
        Next y
                  
    With Columns("H:I")
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
    Selection.ColumnWidth = 22
                  
'
'    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True
'    Application.StatusBar = False

End Sub


Sub StoAddSecondary()    'ready 1

'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
'    Application.StatusBar = "Add Secondary for STO, please wait ......"
    
    Sheets("1.STO").Activate
    Rem add secondary columns for STO sheet
    Rem Range("A2:B66365").ClearContents   '清空A,B列
     Dim k%
     Dim i%
        For k = 1 To 2
            Columns("a:a").Select
            Selection.Insert Shift:=xlToRight
        Next k
            Range("A1") = "'Qty"
            Range("B1") = "'Status"

        For i = 2 To [d66365].End(3).Row
            If Mid(Cells(i, "q"), 1, 1) = "S" And Cells(i, "j") = "Z" Then Cells(i, "b") = "NG" Else Cells(i, "b") = "Avaiable STO"
            If Mid(Cells(i, "q"), 1, 1) = "0" Then Cells(i, "b") = "Picked"
            If Cells(i, "d") <> "" Then Cells(i, "a") = Cells(i, "e") * 1
        Next i
        
    Sheets("2.Yard").Activate
    'add secondary columns for Yard sheet
        Range("S2:S66365").ClearContents   '清空S列
        Range("S1") = "Qty Remaining"
    Dim j%
        For j = 2 To [d66365].End(3).Row
            Cells(j, "s") = (Cells(j, "l") - Cells(j, "m")) * 1
        Next j
'
'    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True
'    Application.StatusBar = False
    
End Sub


Sub TripAvaiableReport()  'ready2

'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
'    Application.StatusBar = "Add Secondary for trip Avaiable Report, please wait ......"
    
    Sheets("4.Trip available").Activate
    'Range("a2:a66365").ClearContents   '清空A列
     Dim k%
        For k = 1 To 1
            Columns("a:a").Select
            Selection.Insert Shift:=xlToRight
        Next k
            Range("A1") = "'Load Date"
            
    Dim d As Object, arr, brr, i&
    Set d = CreateObject("scripting.dictionary")
        d.CompareMode = vbTextCompare '不区分字母大小写
    arr = Sheets("3.Trip report").Range("a2:w66365")
    '数据源装入数组arr
    brr = Sheets("4.Trip available").Range("a2:e66365")
    '查询区域数据装入数组brr
    For i = 1 To UBound(arr)
    '遍历数组arr
        d(arr(i, 1)) = arr(i, 23)
        '将trip 作为key，Load date作为date装入字典
    Next
    For i = 1 To UBound(brr)
    '从brr第一行开始遍历查询数值brr
        If d.exists(brr(i, 5)) Then
        '如果字典中存在trip号
            brr(i, 1) = d(brr(i, 5))
            '根据trip号从字典中取值
        Else
            brr(i, 1) = ""
            '如果字典中不存在相关trip号，则值返回为假空
        End If
    Next
    With Sheets("4.Trip available").Range("a2:e66365")
        .NumberFormat = "@"
        '设置文本格式，避免某些文本数值变形
        .Value = brr
        '结果数组写入单元格区域
    End With
    Set d = Nothing
    '释放字典
    
'    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True
'    Application.StatusBar = False

End Sub



Sub MasterMo()  'ready3
'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
'    Application.StatusBar = "Loading MO, please wait ......"
    
    Sheets("5.MASTER MO").Activate
    'add secondary columns for 5.MASTER MO sheet
        'Range("A2:B66365").ClearContents   '清空A,B列
    Dim k%
        For k = 1 To 2
            Columns("a:a").Select
            Selection.Insert Shift:=xlToRight
        Next k
            Range("A1") = "'Item"
            Range("B1") = "'ETA AT Date"
    
    Dim i%
           For i = 2 To [d66365].End(3).Row
            If Cells(i, "k") = "NS" Then Cells(i, "b") = Cells(i, "c") + 1 Else Cells(i, "b") = Cells(i, "c")
            Cells(i, "a") = "=text(d:d,0)"
            Next i

    '改B栏ETA格式"
    Columns("b:b").Replace What:=" *", Replacement:="", LookAt:=xlPart
    Columns("b:b").TextToColumns Destination:=Range("b1"), _
        DataType:=xlDelimited, FieldInfo:=Array(1, xlMDYFormat)
    Columns("b:b").NumberFormat = "mm/dd/yyyy"
   
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
End Sub




Sub TripAPaste()   'ready 4

    Application.ScreenUpdating = False
    Application.StatusBar = "Trip Fulfill Data Paste, please wait ......"
  
    Sheets("6.Trip Fulfill Data").Select
    Range("K2:AE66365").ClearContents   '清空

    Set CNN = CreateObject("adodb.connection")
    CNN.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=Excel 12.0;Data Source=" & ThisWorkbook.FullName
    Sql = "select * from [4.Trip available$] order by [Item Number] desc,[Load Date]"
    Sheet8.Range("K2").CopyFromRecordset CNN.Execute(Sql)
        
    CNN.Close
    Set CNN = Nothing

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False


End Sub


Sub ColumnAF()   'ready5
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Stock Allocating, please wait ......"

    Sheets("6.Trip Fulfill Data").Activate
    Range("A2:F66365").ClearContents   '清空G,H列
    Columns("B:B").NumberFormat = "#,##0"
    Columns("D:D").NumberFormat = "#,##0"
    
    Dim i%
    For i = 2 To [N66365].End(3).Row
        If Cells(i, "n") = Cells(i - 1, "n") Then Cells(i, "a") = ""
        If Cells(i, "n") <> Cells(i - 1, "n") Then Cells(i, "a").Formula = "=SUMIFS('1.STO'!A:A,'1.STO'!D:D,'6.Trip Fulfill Data'!n:n,'1.STO'!B:B,""=Avaiable STO"")"
        If Cells(i, "n") = Cells(i - 1, "n") Then Cells(i, "b") = ""
        If Cells(i, "n") <> Cells(i - 1, "n") Then Cells(i, "b").Formula = "=SUMIFS('2.Yard'!S:S,'2.Yard'!K:K,'6.Trip Fulfill Data'!n:n)"
        Cells(i, "c") = Cells(i, "q") - Cells(i, "r")
        If Cells(i, "n") <> Cells(i - 1, "n") Then Cells(i, "d") = Cells(i, "a") + Cells(i, "b") - Cells(i, "c") Else Cells(i, "d") = Cells(i - 1, "d") - Cells(i, "c")
        If Cells(i, "c") = 0 Or Cells(i, "d") >= 0 Then Cells(i, "e") = "OK" Else Cells(i, "e") = "Shortage"
        If Cells(i, "e") = "OK" Then Cells(i, "f") = 0
        If Cells(i, "e") <> "OK" And Abs(Cells(i, "d")) < Cells(i, "c") Then Cells(i, "f") = Abs(Cells(i, "d"))
        If Cells(i, "e") <> "OK" And Abs(Cells(i, "d")) >= Cells(i, "c") Then Cells(i, "f") = Abs(Cells(i, "c"))
                
    Next i
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
End Sub

Sub ColumnGH()  'ready6
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Loading Vendor, please wait ......"
    Sheets("6.Trip Fulfill Data").Activate
    Range("G2:H66365").ClearContents   '清空G,H列
    
    Dim m%
        For m = 2 To [N66365].End(3).Row
            Cells(m, "G") = "=VLOOKUP(n:n,Vendor!A:D,4,0)"
            Cells(m, "H") = "=VLOOKUP(n:n,Vendor!A:D,3,0)"

        Next m

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
   

End Sub

Sub ColumnIJK()  'Ready7

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Calculating Shortage Cubes, please wait ......"
    Sheets("6.Trip Fulfill Data").Activate
    'Range("i2:j66365").ClearContents   '清空j列
            
    Dim d As Object, arr, brr, i&
    Set d = CreateObject("scripting.dictionary")
        d.CompareMode = vbTextCompare '不区分字母大小写
    arr = Sheets("Cubes").Range("a2:d108888")
    '数据源装入数组arr
    brr = Sheets("6.Trip Fulfill Data").Range("j2:n66365")
    '查询区域数据装入数组brr
    For i = 1 To UBound(arr)
    '遍历数组arr
        d(arr(i, 1)) = arr(i, 4)
        '将item# 作为key，cubes作为unit cube装入字典
    Next
    For i = 1 To UBound(brr)
    '从brr第一行开始遍历查询数值brr
        If d.exists(brr(i, 5)) Then
        '如果字典中存在trip号
            brr(i, 1) = d(brr(i, 5))
            '根据trip号从字典中取值
        Else
            brr(i, 1) = 0
            '如果字典中不存在相关trip号，则值返回为0
        End If
    Next
    With Sheets("6.Trip Fulfill Data").Range("j2:n66365")
        .NumberFormat = "@"
        '设置文本格式，避免某些文本数值变形
        .Value = brr
        '结果数组写入单元格区域
    End With
    Set d = Nothing
    '释放字典
    
    Dim k%
        For k = 2 To [N66365].End(3).Row
            Cells(k, "i") = Cells(k, "j") * Cells(k, "f")
        Next k
        
    '改K栏日期格式
    Columns("k:k").Replace What:=" *", Replacement:="", LookAt:=xlPart
    Columns("k:k").TextToColumns Destination:=Range("k1"), _
    DataType:=xlDelimited, FieldInfo:=Array(1, xlMDYFormat)
    Columns("k:k").NumberFormat = "mm/dd/yyyy"
    Range("k1") = "Load Date"
        
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False


End Sub



Sub TripReadyList()   'ready 9

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Calculating Trip Ready List, please wait ......"
    
    Sheet8.Activate
    Sheet8.Range("c1") = "Trip_Demand"
    Sheet8.Range("f1") = "Shortage_Piece"
    Sheet8.Range("k1") = "Load_Date"
    Sheet8.Range("m1") = "Dispatch_Date"
    Sheet8.Range("n1") = "Item_Number"
    Sheet8.Range("o1") = "Trip_Number"

    Sheet14.Activate
    Set CNN = CreateObject("adodb.connection")
    CNN.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=Excel 12.0;Data Source=" & ThisWorkbook.FullName

    Dim rs As ADODB.Recordset
    Dim ws As Worksheet
    Dim i As Integer

    Set ws = ThisWorkbook.Worksheets("7. Ready trips")
    ws.Cells.ClearContents
    Sheet14.Range("a1") = "Trip_Demand"
    Sheet14.Range("b1") = "Load_Date"
    Sheet14.Range("c1") = "Dispatch_Date"
    Sheet14.Range("d1") = "Shortage_Piece"
    Sheet14.Range("e1") = "Trip_Demand"

    
    ws.Range("a1:e1") = Array("Trip_Number", "Load_Date", "Dispatch_Date", "Shortage_Piece", "Trip_Demand")
   
    '获取不重复的trip number list
    Sql = "Select Trip_Number,Load_Date,Dispatch_Date,sum(shortage_Piece), sum(Trip_Demand) from [6.Trip Fulfill Data$] Group by Trip_Number,Load_Date,Dispatch_Date having sum(Shortage_Piece)=0 order by Load_Date"
    Set rs = New ADODB.Recordset
    rs.Open Sql, CNN, adOpenKeyset, adLockOptimistic
    ws.Range("A2").CopyFromRecordset rs
    rs.Close
    CNN.Close
    Set rs = Nothing
    Set CNN = Nothing
    
    
    Columns("B:B").NumberFormat = "m/d/yyyy"
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False


End Sub
Sub shortagelist()   'ready 10
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Calculating Shortage List, Please wait ......"
  
    Sheet10.Activate
    Range("A2:E66365").ClearContents   '清空G,H列

    Set CNN = CreateObject("adodb.connection")
    CNN.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=Excel 12.0;Data Source=" & ThisWorkbook.FullName
    Sql = "select [Trip_Number],[Load_Date],[Dispatch_Date],[Item_Number],[Vendor],[Shortage_Piece],[Shortage_Cubes] from [6.Trip Fulfill Data$] where [Shortage_Piece]>0 order by [Item_Number],[Load_Date],[Dispatch_Date]"
    Sheet10.Range("a2").CopyFromRecordset CNN.Execute(Sql)
        
    CNN.Close
    Set CNN = Nothing

    Columns("B:B").NumberFormat = "m/d/yyyy"
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False

End Sub

Sub Copy_CGPO_UPHMO()   'ready 11

        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.StatusBar = "Loading PO and MO, please wait ......"

        Application.ScreenUpdating = False
        Worksheets("9.Supply").Select
        Range("A2:E66365").ClearContents   '清空G,H列
        
        Worksheets("5.MASTER MO").Select
        Range("A2", Range("A2").End(xlDown)).Copy
    
        Worksheets("9.Supply").Range("A2").PasteSpecial Paste:=xlPasteValues    'Copy MO item# 贴上值
        
        Range("H2", Range("H2").End(xlDown)).Copy
        Worksheets("9.Supply").Range("B2").PasteSpecial Paste:=xlPasteValues    'Copy MO Allocated Qty 贴上值
        
        Range("B2", Range("B2").End(xlDown)).Copy
        Worksheets("9.Supply").Range("C2").PasteSpecial Paste:=xlPasteValues    'Copy MO date 贴上值


        Worksheets("OPEN PO").Select
        Range("B2:C2", Range("B2:C2").End(xlDown)).Copy
        Worksheets("9.Supply").Range("A2").End(xlDown).Offset(1, 0).PasteSpecial Paste:=xlPasteValues    'paste PO item 与数量
        Range("A2", Range("A2").End(xlDown)).Copy
        Worksheets("9.Supply").Range("C2").End(xlDown).Offset(1, 0).PasteSpecial Paste:=xlPasteValues   'paste PO due date

        Sheet15.Select
        Range("a1") = "Item_Number"
        Range("b1") = "Qty"
        Range("c1") = "ETA_Date"
        

        '改K栏日期格式
        Columns("c:c").Replace What:=" *", Replacement:="", LookAt:=xlPart
        Columns("c:c").TextToColumns Destination:=Range("c1"), _
        DataType:=xlDelimited, FieldInfo:=Array(1, xlMDYFormat)
        Columns("c:c").NumberFormat = "mm/dd/yyyy"
    
    
        Range("A1").CurrentRegion.Sort key1:=Range("A2"), order1:=xlAscending, key2:=Range("C2"), order2:=xlAscending, Header:=xlYes
        Rem 按款号与日期排序
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False

End Sub


Sub Qiliao()   'ready 12

        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.StatusBar = " Allocation by PO and MO, please wait ......"


Dim arr1(), arr2()
Dim a%, b%

Sheets("8.Shortage trips").Select
'Range("b1", Cells(Rows.Count, "g").End(xlUp)).Sort [d1], xlAscending, [b1], , xlAscending, , , xlYes  '按款号与日期排序
arr1 = Range("a2", Cells(Cells(Rows.Count, "g").End(xlUp).Row, "h"))    '数组赋值
Sheets("9.Supply").Select
'Range("a1", Cells(Rows.Count, "c").End(xlUp)).Sort [a1], xlAscending, [c1], , xlAscending, , , xlYes   '按款号与日期排序
arr2 = Range("a2", Cells(Rows.Count, "c").End(xlUp)) '数组赋值

'For a = LBound(arr1) To UBound(arr1)
'    arr1(a, 6) = arr1(a, 6) * 1   '将文字改为数值
'Next

Sheets("8.Shortage trips").Select
Range("h2:h66365").ClearContents

100
     For a = LBound(arr1) To UBound(arr1)
            For b = LBound(arr2) To UBound(arr2)
                If arr1(a, 4) = arr2(b, 1) And arr1(a, 6) <= arr2(b, 2) And arr1(a, 6) <> 0 Then  '订单款号与PO款号相同，且 订单需求<=PO数量 且 订单需求不为0
                    arr2(b, 2) = arr2(b, 2) - arr1(a, 6)                'PO数量 = PO数量 - 订单需求
                    arr1(a, 6) = arr1(a, 6) - arr1(a, 6)                '订单需求  = 订单需求 - 订单需求
                    Sheets("8.Shortage trips").Cells(a + 1, "h") = arr2(b, 3)            '预计齐料日 = PO的预计进料日期   ‘cells(a+1)是为了跟数组第一行对齐
                    arr1(a, 8) = arr2(b, 3)
                ElseIf arr1(a, 4) = arr2(b, 1) And arr1(a, 6) > arr2(b, 2) And arr2(b, 2) <> 0 Then  '亦或，订单款号与PO款号相同，且 订单需求>PO数量 且 订单需求不为0
                    arr1(a, 6) = arr1(a, 6) - arr2(b, 2)       '订单缺货数量 = 订单缺货数量 - PO数量
                    arr2(b, 2) = arr2(b, 2) - arr2(b, 2)       'PO数量 = PO数量 - PO数量
                    If arr1(a, 6) <> 0 Then                    '如果订单需求<>0, 回到100
                    GoTo 100
                    End If
                End If
            Next
        Next
        


Sheets("8.Shortage trips").Select

Dim i%
    For i = 2 To [h66365].End(3).Row
        If Cells(i, "h") = "" Then Cells(i, "h") = #12/31/2099#
    Next


    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False


End Sub

  
Sub EstimatedReadyTrips()  'ready 13

        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.StatusBar = "Calculating Estimate Ready Trips, please waiting......"
        Sheets("10.Estimated ready trips").Select
        Range("a2:e66365").ClearContents
                
        Dim CNN As Object
        Set CNN = CreateObject("adodb.connection")
        CNN.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties='excel 12.0;hdr=yes;imex=1';Data Source=" & ThisWorkbook.FullName
        
        S1 = "select Trip_Number,Load_Date,Dispatch_Date,max(ETA) as Estimated_Ready_Date from [8.Shortage trips$] Group by Trip_Number,Load_Date,Dispatch_Date Order by Load_Date"
        Sheets("10.Estimated ready trips").[a2].CopyFromRecordset CNN.Execute(S1)
        
        CNN.Close
        Set CNN = Nothing
        
        Sheets("10.Estimated ready trips").Select
        Range("a1") = "Trip_Number"
        Range("b1") = "Load_Date"
        Range("c1") = "Dispatch_Date"
        Range("d1") = "Estimated_Ready_Date"
        
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False



End Sub

Sub Cubes()  'ready 5.1

    Dim i As Long
    Dim adors As New Recordset
    
    Worksheets("Cubes").Range("A2:D66365").ClearContents
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Loading Unit Cubes, please wait ......"
    
    Set Db = New Connection
    Db.CursorLocation = adUseClient
    If Db.State = 1 Then Db.Close
    
    UName = Sheet11.Range("a1")
    UPass = Sheet11.Range("a2")
    DateStart = Sheet11.Range("a3")
    DateEnd = Sheet11.Range("a4")
    
   
    Db.Open "Provider =IBMDASQL.DataSource.1" & _
     ";Catalog Library List=JIMTDTA" & _
     ";Persist Security Info=True" & _
     ";Force Translate=0" & _
     ";Data Source = WVFHA" & _
     ";User ID = " & UName & "" & _
     ";Password = " & UPass
     
     Set adors = New Recordset
     If adors.State = 1 Then adors.Close

    cmdtxt = "SELECT ITMRVA.ITNBR, ITEMBL.HOUSE, ITEMBL.ITCLS, ITMRVA.B2Z95S as UnitCube " & _
             "FROM D20ACF9V.AFILELIB.ITBEXT ITBEXT, D20ACF9V.AMFLIBA.ITEMBL ITEMBL, D20ACF9V.AFILELIB.ITMEXT ITMEXT, D20ACF9V.AMFLIBA.ITMRVA ITMRVA, D20ACF9V.AMFLIBA.WHSMST WHSMST " & _
             "WHERE ITMRVA.ITNBR = ITEMBL.ITNBR AND ITMRVA.STID = WHSMST.STID AND ITBEXT.HOUSE = ITEMBL.HOUSE AND ITBEXT.ITNBR = ITEMBL.ITNBR AND ITBEXT.ITNBR = ITMRVA.ITNBR AND ITMEXT.ITNBR = ITBEXT.ITNBR AND ITMEXT.ITNBR = ITEMBL.ITNBR AND ITMEXT.ITNBR = ITMRVA.ITNBR AND ITEMBL.HOUSE = WHSMST.WHID AND (ITEMBL.HOUSE='335') and ITEMBL.ITCLS LIKE 'Z%' AND ITEMBL.ITCLS NOT LIKE '%K' " & _
             "ORDER BY ITMRVA.ITNBR"
        
'SELECT ITMRVA.ITNBR, ITEMBL.HOUSE, ITEMBL.ITCLS, ITMRVA.B2Z95S as UnitCube
'FROM D20ACF9V.AFILELIB.ITBEXT ITBEXT, D20ACF9V.AMFLIBA.ITEMBL ITEMBL, D20ACF9V.AFILELIB.ITMEXT ITMEXT, D20ACF9V.AMFLIBA.ITMRVA ITMRVA, D20ACF9V.AMFLIBA.WHSMST WHSMST
'WHERE ITMRVA.ITNBR = ITEMBL.ITNBR AND ITMRVA.STID = WHSMST.STID AND ITBEXT.HOUSE = ITEMBL.HOUSE AND ITBEXT.ITNBR = ITEMBL.ITNBR AND ITBEXT.ITNBR = ITMRVA.ITNBR AND ITMEXT.ITNBR = ITBEXT.ITNBR AND ITMEXT.ITNBR = ITEMBL.ITNBR AND ITMEXT.ITNBR = ITMRVA.ITNBR AND ITEMBL.HOUSE = WHSMST.WHID AND (ITEMBL.HOUSE='335') and ITEMBL.ITCLS LIKE 'Z%' AND ITEMBL.ITCLS NOT LIKE '%K'
'ORDER BY ITMRVA.ITNBR

    adors.Open cmdtxt, Db, 3, 3
     For i = 0 To adors.Fields.Count - 1
         Worksheets("Cubes").Cells(1, i + 1) = adors.Fields(i).Name
     Next i
     
     Worksheets("Cubes").Range("A2").CopyFromRecordset adors
     adors.Close
     Set adors = Nothing

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False



End Sub

Sub OpenPO()  'ready 5.2

    Dim i As Long
    Dim adors As New Recordset
    
    Worksheets("OPEN PO").Range("A2:O66365").ClearContents
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Loading Open PO, please wait ......"
    
    Set Db = New Connection
    Db.CursorLocation = adUseClient
    If Db.State = 1 Then Db.Close
    
    UName = Sheet11.Range("a1")
    UPass = Sheet11.Range("a2")
    DateStart = Sheet11.Range("a3")
    DateEnd = Sheet11.Range("a4")
    
   
    Db.Open "Provider =IBMDASQL.DataSource.1" & _
     ";Catalog Library List=JIMTDTA" & _
     ";Persist Security Info=True" & _
     ";Force Translate=0" & _
     ";Data Source = WVFHA" & _
     ";User ID = " & UName & "" & _
     ";Password = " & UPass
     
     Set adors = New Recordset
     If adors.State = 1 Then adors.Close

    cmdtxt = "SELECT POITEM.ITNBR, POITEM.QTYOR, POMAST.HOUSE, POMAST.ORDNO, POMAST.VNDNR, POMAST.PSTTS,ITMEXT.UUCCIM, POITEM.DUEDT, ITBEXT.ITMCLSID, ITBEXT.PICKPUT, ITBEXT.SCOOPQTY, ITBEXT.SKIDSIZE,VENNAML0.VNNMVM " & _
             "FROM AFILELIB.ITBEXT ITBEXT, AFILELIB.ITMEXT ITMEXT, AMFLIBA.POITEM POITEM, AMFLIBA.POMAST POMAST, AMFLIBA.VENNAML0 VENNAML0 " & _
             "WHERE ITBEXT.ITNBR = ITMEXT.ITNBR AND POITEM.ITNBR = ITMEXT.ITNBR AND POITEM.ORDNO = POMAST.ORDNO AND POMAST.VNDNR = VENNAML0.VNDRVM AND POITEM.HOUSE = ITBEXT.HOUSE AND POMAST.HOUSE = POITEM.HOUSE AND ((POMAST.PSTTS='10') OR (POMAST.PSTTS='20') OR (POMAST.PSTTS='30')) AND (POITEM.HOUSE='335') AND (POMAST.VNDNR NOT IN ('600039','900639','900515')) " & _
             "ORDER BY POITEM.ITNBR, POITEM.DUEDT"
'
'SELECT POITEM.ITNBR, POITEM.QTYOR, POMAST.HOUSE, POMAST.ORDNO, POMAST.VNDNR, POMAST.PSTTS,ITMEXT.UUCCIM, POITEM.DUEDT, ITBEXT.ITMCLSID, ITBEXT.PICKPUT, ITBEXT.SCOOPQTY, ITBEXT.SKIDSIZE,VENNAML0.VNNMVM
'FROM AFILELIB.ITBEXT ITBEXT, AFILELIB.ITMEXT ITMEXT, AMFLIBA.POITEM POITEM, AMFLIBA.POMAST POMAST, AMFLIBA.VENNAML0 VENNAML0
'WHERE ITBEXT.ITNBR = ITMEXT.ITNBR AND POITEM.ITNBR = ITMEXT.ITNBR AND POITEM.ORDNO = POMAST.ORDNO AND POMAST.VNDNR = VENNAML0.VNDRVM AND POITEM.HOUSE = ITBEXT.HOUSE AND POMAST.HOUSE = POITEM.HOUSE AND ((POMAST.PSTTS='10') OR (POMAST.PSTTS='20') OR (POMAST.PSTTS='30')) AND (POITEM.HOUSE='335') AND (POMAST.VNDNR NOT IN ('600039','900639','900515'))
'ORDER BY POITEM.ITNBR, POITEM.DUEDT

    adors.Open cmdtxt, Db, 3, 3
     For i = 0 To adors.Fields.Count - 1
         Worksheets("Open PO").Cells(1, i + 2) = adors.Fields(i).Name
     Next i
     
     Worksheets("Open PO").Range("B2").CopyFromRecordset adors
     adors.Close
     Set adors = Nothing

    
    Sheets("OPEN PO").Activate
    Range("A1") = "Due_Date"
    
    Dim m%
        For m = 2 To [B66365].End(3).Row
            Cells(m, "a").Formula = "=date(""20""&mid(i:i,2,2),mid(i:i,4,2),mid(i:i,6,2))"
            
        Next m

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False


End Sub

Sub LoadPO()

'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
'    Application.StatusBar = "Loading OpenPO, please wait ......"


    Dim i As Long
    Dim adors As New Recordset
    
    
    
    Worksheets("OpenPO").Range("A2:y66365").ClearContents
    
    
    Set Db = New Connection
    Db.CursorLocation = adUseClient
    If Db.State = 1 Then Db.Close
    
    UName = Sheet11.Range("a1")
    UPass = Sheet11.Range("a2")
    DateStart = Sheet11.Range("a3")
    DateEnd = Sheet11.Range("a4")
    
   
    Db.Open "Provider =IBMDASQL.DataSource.1" & _
     ";Catalog Library List=JIMTDTA" & _
     ";Persist Security Info=True" & _
     ";Force Translate=0" & _
     ";Data Source = WVFHA" & _
     ";User ID = " & UName & "" & _
     ";Password = " & UPass
     
     Set adors = New Recordset
     If adors.State = 1 Then adors.Close

    cmdtxt = "SELECT POITEM.ITNBR, POITEM.QTYOR, POMAST.HOUSE, POMAST.ORDNO, POMAST.VNDNR, POMAST.PSTTS, POITEM.STKQT, POITEM.STKDT, POITEM.DYLDE, POITEM.DYLDL, ITMEXT.UUCCIM, POITEM.DUEDT, ITBEXT.ITMCLSID, ITBEXT.PICKPUT, ITBEXT.SCOOPQTY, ITBEXT.SKIDSIZE, POITEM.DOKDT, POITEM.MSNDD, POITEM.MSNSD, VENNAML0.VNNMVM, POITEM.QTYOR, POITEM.QTDEV, POITEM.STKQT,ITEMBL.ITCLS " & _
             "FROM AFILELIBQ.ITBEXT ITBEXT, AFILELIBQ.ITMEXT ITMEXT, AMFLIBQ.POITEM POITEM, AMFLIBQ.POMAST POMAST, AMFLIBQ.VENNAML0 VENNAML0,AMFLIBQ.ITEMBL ITEMBL " & _
             "WHERE ITBEXT.ITNBR = ITMEXT.ITNBR AND POITEM.ITNBR = ITMEXT.ITNBR AND ITMEXT.ITNBR=ITEMBL.ITNBR AND POITEM.ORDNO = POMAST.ORDNO AND POMAST.VNDNR = VENNAML0.VNDRVM AND POITEM.HOUSE = ITBEXT.HOUSE AND POMAST.HOUSE = POITEM.HOUSE AND POITEM.HOUSE = ITEMBL.HOUSE AND ((POMAST.PSTTS='10') OR (POMAST.PSTTS='20') OR (POMAST.PSTTS='30')) AND (POITEM.HOUSE='232') and (ITEMBL.ITCLS not like 'Z%') " & _
             "ORDER BY POITEM.ITNBR, POITEM.DUEDT"


'SELECT POITEM.ITNBR, POITEM.QTYOR, POMAST.HOUSE, POMAST.ORDNO, POMAST.VNDNR, POMAST.PSTTS, POITEM.STKQT, POITEM.STKDT, POITEM.DYLDE, POITEM.DYLDL, ITMEXT.UUCCIM, POITEM.DUEDT, ITBEXT.ITMCLSID, ITBEXT.PICKPUT, ITBEXT.SCOOPQTY, ITBEXT.SKIDSIZE, POITEM.DOKDT, POITEM.MSNDD, POITEM.MSNSD, VENNAML0.VNNMVM, POITEM.QTYOR, POITEM.QTDEV, POITEM.STKQT,ITEMBL.ITCLS
'FROM AFILELIB.ITBEXT ITBEXT, AFILELIB.ITMEXT ITMEXT, AMFLIBA.POITEM POITEM, AMFLIBA.POMAST POMAST, AMFLIBA.VENNAML0 VENNAML0,AMFLIBA.ITEMBL ITEMBL
'WHERE ITBEXT.ITNBR = ITMEXT.ITNBR AND POITEM.ITNBR = ITMEXT.ITNBR AND ITMEXT.ITNBR=ITEMBL.ITNBR AND POITEM.ORDNO = POMAST.ORDNO AND POMAST.VNDNR = VENNAML0.VNDRVM AND POITEM.HOUSE = ITBEXT.HOUSE AND POMAST.HOUSE = POITEM.HOUSE AND POITEM.HOUSE = ITEMBL.HOUSE AND ((POMAST.PSTTS='10') OR (POMAST.PSTTS='20') OR (POMAST.PSTTS='30')) AND (POITEM.HOUSE='335') and (ITEMBL.ITCLS not like 'Z%')
'
    adors.Open cmdtxt, Db, 3, 3
     For i = 0 To adors.Fields.Count - 1
         Worksheets("OpenPO").Cells(1, i + 2) = adors.Fields(i).Name
     Next i
     
     Worksheets("OpenPO").Range("B2").CopyFromRecordset adors
     adors.Close
     Set adors = Nothing

    
    Sheets("OpenPO").Activate
    Range("A1") = "Due_Date"
    
    Dim m%
        For m = 2 To [B66365].End(3).Row
            Cells(m, "a").Formula = "=date(""20""&mid(m:m,2,2),mid(m:m,4,2),mid(m:m,6,2))"
            
        Next m

'    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True
'    Application.StatusBar = False


End Sub

Sub RpOpenOrderLoading()


'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
'    Application.StatusBar = "Loading RP Open orders, please wait ......"

    Dim i As Long
    Dim adors As New Recordset

    Worksheets("RPOpenOrders").Range("A2:W66365").ClearContents

    
    Set Db = New Connection
    Db.CursorLocation = adUseClient
    If Db.State = 1 Then Db.Close
    
    UName = Sheet11.Range("a1")
    UPass = Sheet11.Range("a2")
    DateStart = Sheet11.Range("a3")
    DateEnd = Sheet11.Range("a4")
    
   
    Db.Open "Provider =IBMDASQL.DataSource.1" & _
     ";Catalog Library List=JIMTDTA" & _
     ";Persist Security Info=True" & _
     ";Force Translate=0" & _
     ";Data Source = WVFHA" & _
     ";User ID = " & UName & "" & _
     ";Password = " & UPass
     
     Set adors = New Recordset
     If adors.State = 1 Then adors.Close

    cmdtxt = "SELECT ARPHEDR.CUSNO, CUSMAS.CUSNM, ARPHEDR.ENTDAT, ARPHEDR.RPKEY, ARPHEDR.MODEL, ARPHEDR.SHPDAT, ARPHEDR.ORDPCK, ARPHEDR.TRIP#, ARPHEDR.WHOSE, ARPHEDR.SHPCTY, ARPDETL.ITEMNO, ITMRVA.ITDSC, ARPDETL.QTY, ARPDETL.SHPFLG, ARPDETL.PCKDTE, ARPDETL.PCKTME " & _
             "FROM AFILELIBQ.ARPDETL ARPDETL, AFILELIBQ.ARPHEDR ARPHEDR, AMFLIBQ.CUSMAS CUSMAS, AMFLIBQ.ITMRVA ITMRVA " & _
             "WHERE ARPHEDR.RPKEY = ARPDETL.RPKEY AND ARPHEDR.CUSNO = CUSMAS.CUSNO AND ARPDETL.ITEMNO = ITMRVA.ITNBR AND ARPHEDR.WHOSE = ITMRVA.STID AND ((ARPHEDR.ACTCOD='A') AND (ARPHEDR.ENTDAT Between 20190101 And 20221231 ) AND (ARPHEDR.WHOSE='232') AND (ARPHEDR.SHPDAT=0))" & _
             "ORDER BY ARPHEDR.ENTDAT, ARPHEDR.RPKEY"


    adors.Open cmdtxt, Db, 3, 3
     For i = 0 To adors.Fields.Count - 1
         Worksheets("RPOpenOrders").Cells(1, i + 1) = adors.Fields(i).Name
     Next i
     
     Worksheets("RPOpenOrders").Range("A2").CopyFromRecordset adors
     adors.Close
     Set adors = Nothing

    
'    Sheets("RPOpenOrders").Activate
'    Range("A1") = "Due_Date"
    
'    Dim m%
'        For m = 2 To [B66365].End(3).Row
'            Cells(m, "a").Formula = "=date(""20""&mid(i:i,2,2),mid(i:i,4,2),mid(i:i,6,2))"
'
'        Next m

'    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True
'    Application.StatusBar = False


End Sub

Sub OrderStatusJudge()

'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
'    Application.StatusBar = "Add column Type, please wait ......"
    
    Sheets("RPOpenOrders").Activate
    Dim i%
    For i = 2 To [d66365].End(3).Row
        If Cells(i, "f") <> 0 Then Cells(i, "r") = "Shipped"
        If Cells(i, "f") = 99999999 Then Cells(i, "r") = "Order Cancelled"
        If Cells(i, "h") = 0 And Cells(i, "o") > 0 Then Cells(i, "r") = "Packed and waiting for stick on trip"
        If Cells(i, "h") = 0 And Cells(i, "g") > 0 Then Cells(i, "r") = "Packed and waiting for stick on trip"
        If Cells(i, "h") > 0 And Cells(i, "o") > 0 Then Cells(i, "r") = "Stuck on trip and waiting for picking"
        If Cells(i, "h") > 0 And Cells(i, "g") > 0 Then Cells(i, "r") = "Stuck on trip and waiting for picking"
        If Cells(i, "h") = 0 And Cells(i, "g") = 0 And Cells(i, "o") = 0 And Cells(i, "f") = 0 Then Cells(i, "r") = "Still not pick&pack"
        Next i

'    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True
'    Application.StatusBar = False

End Sub

Sub RPOrderCounted()

'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
'    Application.StatusBar = "Counting orders, please wait ......"
    
     Sheets("RPOpenOrders").Activate
     Range("s2:s66365").ClearContents
     Dim i%, row1%
        For i = 2 To [d66365].End(3).Row
            row1 = Application.WorksheetFunction.CountIf(Range("d1:d" & i), Cells(i, "d"))
            If row1 = 1 Then Cells(i, "s") = 1
            If row1 > 1 Then Cells(i, "s") = 0
            Next i
    
'    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True
'    Application.StatusBar = False

End Sub

Sub rp_creation_date_pending_days()

'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
'    Application.StatusBar = "Changing order creation date, please wait ......"

    Sheets("RPOpenOrders").Activate
    Range("t2:t66365").ClearContents
    
    Dim datetmp As Date
    Dim i%
    For i = 2 To [d66365].End(3).Row
        Cells(i, "t") = Left(Cells(i, "c"), 4) & "/" & Mid(Cells(i, "c"), 5, 2) & "/" & Right(Cells(i, "c"), 2)
        Cells(i, "u") = Date - Cells(i, "t")
    
    Next
    
'    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True
'    Application.StatusBar = False

'Backup code
'datetmp = Format(CLng(strtmp), "0-00-00")
'直接这样也可以:
'datetmp = Format(strtmp, "0-00-00")


End Sub

Sub Column_V_Intervals()
'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
'    Application.StatusBar = "Stock Allocating, please wait ......"

    Sheets("RPOpenOrders").Activate
    Range("v2:v66365").ClearContents   '清空G,H列
    
    Dim i%
    For i = 2 To [u66365].End(3).Row
    
        Cells(i, "v").Value = Application.WorksheetFunction.Lookup(Sheet1.Cells(i, "u").Value, Sheet3.Range("a1:a11"), Sheet3.Range("b1:b11"))
      
        '=LOOKUP(U:U,intervals!$A$1:$A$10,intervals!$B$1:$B$11)
         Next i
    
'    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True
'    Application.StatusBar = False
    
End Sub


Sub intervals_no()

'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
'    Application.StatusBar = "Calculating, please wait ......"

    Dim d As Object, arr, brr, i&
    
    Set d = CreateObject("scripting.dictionary")
        d.CompareMode = vbTextCompare   '不区分大小写

        arr = Sheets("intervals").[a1].CurrentRegion            '数据源装入数组arr
        brr = Sheets("RPOpenOrders").[a1].CurrentRegion    '查询区域数据装入数组brr
    
        For i = 1 To UBound(arr)    '遍历数组arr
            d(arr(i, 2)) = arr(i, 3) '将intervals作为key, No作为item装入字典
            Next
        
        For i = 2 To UBound(brr) '标题行不用查询，所以从2开始遍历查询数值brr
            If d.exists(brr(i, 22)) Then   '如果字典中存在intervals的值
                brr(i, 23) = d(brr(i, 22))    '根据区间从字典中取值
                Else
                brr(i, 23) = "view"    '如果字典中不存在intervals则返回view
             End If
                
                Cells(i, "w") = brr(i, 23)
            Next
        Set d = Nothing
    
    
'    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True
'    Application.StatusBar = False
    
        
        
End Sub



Sub Shortage_vs_PO()


'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
'    Application.StatusBar = "Calculating, please wait ......"

Sheet10.Activate
Range("a1:n66653").ClearContents

Set CNN = CreateObject("adodb.connection")
CNN.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=Excel 12.0;Data Source=" & ThisWorkbook.FullName
Sql = "select [CUSNO],[CUSNM],[RPKEY],[MODEL],[ITEMNO],[ITDSC],[QTY],[RP Order#],[Order Creation Date],[Pending days],[Intervals],[Balance],[Item Status] " _
    & "from [Unpick_orders$] " _
    & "where [Item Status]=""Shortage"" " _
    & "Order by [ITEMNO],[Order Creation Date]"

Sheet10.Range("a2").CopyFromRecordset CNN.Execute(Sql)
CNN.Close
Set CNN = Nothing

Range("a1:m1") = Array("CUSNO", "CUSNM", "RPKEY", "MODEL", "ITEMNO", "ITDSC", "QTY", "RP Order#", "Order Creation Date", "Pending days", "Intervals", "Balance", "Item Status")
Columns("I:I").NumberFormat = "m/d/yyyy"
Columns("A:M").EntireColumn.AutoFit

Range("A1:m1").Select
    Range("m1").Activate
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
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
    Columns("A:m").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
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

Dim i%

    For i = 2 To [M66365].End(3).Row
        Range("N1") = "ShortagePieces"
        If Cells(i, "e") <> Cells(i - 1, "e") And Abs(Cells(i, "l")) <= Cells(i, "g") Then Cells(i, "N") = Abs(Cells(i, "L")) Else Cells(i, "N") = Abs(Cells(i, "G"))
        Next i
                
'    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True
'    Application.StatusBar = False
End Sub




Sub InsertSummaryPivot1()
'Macro By ExcelChamps 增加pivotTable到Summary Sheet


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
Dim i%

'Insert a New Blank Worksheet

Set PSheet = Worksheets("Summary")
Set DSheet = Worksheets("RPOpenOrders")
Worksheets("Summary").Activate
Cells.ClearContents


On Error Resume Next
ActiveSheet.PivotTables("PivotTable2").TableRange2.Clear

'Define Data Range
LastRow = DSheet.Cells(Rows.Count, 1).End(xlUp).Row
LastCol = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column
Set PRange = DSheet.Cells(1, 1).Resize(LastRow, LastCol)


'Define Pivot Cache '确定pivottable的位置 cells(3,1)， 重要：PCache需要Microsoft ActiveX Data Objects 2.6

Set PCache = ActiveWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:=PRange). _
CreatePivotTable(TableDestination:=PSheet.Cells(3, 1), _
TableName:="PivotTable2")

'Insert Blank Pivot Table
Set PTable = PCache.CreatePivotTable _
(TableDestination:=PSheet.Cells(1, 1), TableName:="PivotTable2")

'Insert Row Fields    '增加row值
With ActiveSheet.PivotTables("PivotTable2").PivotFields("NO")
.Orientation = xlRowField
.Position = 1
End With

With ActiveSheet.PivotTables("PivotTable2").PivotFields("Intervals")
.Orientation = xlRowField
.Position = 2
End With

'With ActiveSheet.PivotTables("PivotTable2").PivotFields("Intervals")
'.Orientation = xlRowField
'.Position = 3
'End With

'Insert Column Fields    '增加column值
With ActiveSheet.PivotTables("PivotTable2").PivotFields("Type")
.Orientation = xlColumnField
.Position = 1
End With

'Insert Data Field
With ActiveSheet.PivotTables("PivotTable2").PivotFields("RP Order#")
.Orientation = xlDataField
.Position = 1
.Function = xlSum
.NumberFormat = "#,##0"
End With

'With ActiveSheet.PivotTables("PivotTable1").PivotFields("QTY")
'.Orientation = xlDataField
'.Position = 2
'.Function = xlSum
'.NumberFormat = "#,##0"
'End With

'With ActiveSheet.PivotTables("PivotTable1").PivotFields("QTY")
'.Orientation = xlDataField
'.Position = 3
'.Calculation = xlPercentOfTotal
'.NumberFormat = "0.00%"
'.Caption = "%(QTY)"
'End With

'Format Pivot Table
ActiveSheet.PivotTables("PivotTable2").ShowTableStyleRowStripes = True
ActiveSheet.PivotTables("PivotTable2").TableStyle2 = "PivotStyleMedium9"
ActiveSheet.PivotTables("PivotTable2").MergeLabels = True  '

''经典格式
With ActiveSheet.PivotTables("PivotTable2")
    .InGridDropZones = True
    .RowAxisLayout xlTabularRow
End With

With ActiveSheet.PivotTables("PivotTable2")
    .PivotFields("NO").Subtotals(1) = False
    .PivotFields("Intervals").Subtotals(1) = False
End With

'置中对齐
Columns("C:F").Select
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
Columns("C:C").ColumnWidth = 17.86
Columns("D:D").ColumnWidth = 12.14
Columns("E:E").ColumnWidth = 17.43
Columns("F:F").ColumnWidth = 10.14
ActiveSheet.PivotTables("PivotTable2").HasAutoFormat = False



    For i = 5 To 19
        
        If Cells(i, "F") <> "" Then Cells(i, "G") = Cells(i, "F") / Cells(Cells(5, "F").End(4).Row, "f")
        
         Cells(i, "g").NumberFormat = "0.0%"
        Cells(4, "G") = "%"
        Next
        
        
'    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True
'    Application.StatusBar = False

End Sub



Sub InsertSummaryPivot2()
'Macro By ExcelChamps 增加pivotTable到Summary Sheet

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
Dim i%
'Insert a New Blank Worksheet

Set PSheet = Worksheets("Summary")
Set DSheet = Worksheets("RPOpenOrders")
Worksheets("Summary").Activate
On Error Resume Next
ActiveSheet.PivotTables("PivotTable3").TableRange2.Clear

'Define Data Range
LastRow = DSheet.Cells(Rows.Count, 1).End(xlUp).Row
LastCol = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column
Set PRange = DSheet.Cells(1, 1).Resize(LastRow, LastCol)


'Define Pivot Cache '确定pivottable的位置 cells(20,1)
Set PCache = ActiveWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:=PRange). _
CreatePivotTable(TableDestination:=PSheet.Cells(20, 1), _
TableName:="PivotTable3")

'Insert Blank Pivot Table
Set PTable = PCache.CreatePivotTable _
(TableDestination:=PSheet.Cells(18, 1), TableName:="PivotTable3")

'Insert Row Fields    '增加row值
With ActiveSheet.PivotTables("PivotTable3").PivotFields("NO")
.Orientation = xlRowField
.Position = 1
End With

With ActiveSheet.PivotTables("PivotTable3").PivotFields("Intervals")
.Orientation = xlRowField
.Position = 2
End With

'With ActiveSheet.PivotTables("PivotTable2").PivotFields("Intervals")
'.Orientation = xlRowField
'.Position = 3
'End With

'Insert Column Fields    '增加column值
With ActiveSheet.PivotTables("PivotTable3").PivotFields("Type")
.Orientation = xlColumnField
.Position = 1
End With

'Insert Data Field
With ActiveSheet.PivotTables("PivotTable3").PivotFields("QTY")
.Orientation = xlDataField
.Position = 1
.Function = xlSum
.NumberFormat = "#,##0"
End With

'With ActiveSheet.PivotTables("PivotTable1").PivotFields("QTY")
'.Orientation = xlDataField
'.Position = 2
'.Function = xlSum
'.NumberFormat = "#,##0"
'End With

'With ActiveSheet.PivotTables("PivotTable1").PivotFields("QTY")
'.Orientation = xlDataField
'.Position = 3
'.Calculation = xlPercentOfTotal
'.NumberFormat = "0.00%"
'.Caption = "%(QTY)"
'End With

'Format Pivot Table
ActiveSheet.PivotTables("PivotTable3").ShowTableStyleRowStripes = True
ActiveSheet.PivotTables("PivotTable3").TableStyle2 = "PivotStyleMedium9"
ActiveSheet.PivotTables("PivotTable3").MergeLabels = True  '

''经典格式
With ActiveSheet.PivotTables("PivotTable3")
    .InGridDropZones = True
    .RowAxisLayout xlTabularRow
End With

With ActiveSheet.PivotTables("PivotTable3")
    .PivotFields("NO").Subtotals(1) = False
    .PivotFields("Intervals").Subtotals(1) = False
End With

'置中对齐
Columns("C:F").Select
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
Columns("C:C").ColumnWidth = 17.86
Columns("D:D").ColumnWidth = 12.14
Columns("E:E").ColumnWidth = 17.43
Columns("F:F").ColumnWidth = 10.14
Rows("4:4").RowHeight = 35.25
Rows("19:19").RowHeight = 35.25
ActiveSheet.PivotTables("PivotTable3").HasAutoFormat = False

    For i = 22 To 38
        
        If Cells(i, "F") <> "" Then Cells(i, "G") = Cells(i, "F") / Cells(Cells(21, "F").End(4).Row, "f")
        
         Cells(i, "g").NumberFormat = "0.0%"
        Cells(21, "G") = "%"
        Next
        
'    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True
'    Application.StatusBar = False

End Sub





Sub unfilter2()
    '取消筛选
    
    Dim i%, sht As Worksheet
    
    For Each sht In Worksheets
        ' 如果当前工作表为筛选模式，则取消
        If sht.AutoFilterMode = True Then sht.AutoFilterMode = 0
        ' 如果当前工作表没有筛选，则加上筛选
        If sht.AutoFilterMode = False Then sht.Range("a1").AutoFilter Field:=1
    Next
    
End Sub


Sub loading_unpick_orders()

'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
'    Application.StatusBar = "Calculating, please wait ......"


Sheet5.Activate
Range("a2:ab66653").ClearContents

Set CNN = CreateObject("adodb.connection")
CNN.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=Excel 12.0;Data Source=" & ThisWorkbook.FullName
Sql = "select * from [RPOpenOrders$] where [Type]=""Still not pick&pack"" order by [ITEMNO],[ENTDAT]"
Sheet5.Range("a2").CopyFromRecordset CNN.Execute(Sql)
CNN.Close
Set CNN = Nothing

Columns("t:t").NumberFormat = "m/d/yyyy"

'    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True
'    Application.StatusBar = False


End Sub

Sub ColumnX()  'AVAILABLE STO

'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
'    Application.StatusBar = "Calculating, please wait ......"


Sheet5.Activate
Range("x2:x66653").ClearContents

Dim i%, myv As Variant, arr As Variant, brr


    For i = 2 To [k66365].End(3).Row
    Set arr = Sheet4.Range("i:i")
    Set brr = Sheet4.Range("a:a")
        myv = Application.SumIfs(arr, brr, Cells(i, "k"))
    
    If IsError(myv) Then Cells(i, "x") = 0 Else Cells(i, "x").Value = myv
    If Cells(i, "k") = Cells(i - 1, "k") Then Cells(i, "x").Value = 0
    Cells(i, "x") = Cells(i, "x") * 1
    Next i
 
    With Columns("x:ac")
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
    Selection.ColumnWidth = 22
    
    
'    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True
'    Application.StatusBar = False

End Sub


Sub ColumnYZAA()  'BALANCE CALCULATION
'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
'    Application.StatusBar = "Calculating, please wait ......"


Sheet5.Activate
Range("y2:AA66653").ClearContents

Dim i%

    For i = 2 To [k66365].End(3).Row
    
    If Cells(i, "k") = Cells(i - 1, "k") Then Cells(i, "y") = Cells(i - 1, "y") - Cells(i, "m")
    If Cells(i, "k") <> Cells(i - 1, "k") Then Cells(i, "y") = Cells(i, "x") - Cells(i, "m")
    If Cells(i, "y") < 0 Then Cells(i, "z") = "Shortage" Else Cells(i, "z") = "OK"
    If Cells(i, "z") = "OK" Then Cells(i, "AA") = 0 Else Cells(i, "AA") = 1
    
    Next i
    
    
'    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True
'    Application.StatusBar = False
End Sub


Sub Order_fulfillment()  'Order Avaiable STO Fulfillment

'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
'    Application.StatusBar = "Calculating, please wait ......"

Sheet5.Activate
Range("AB2:AB66653").ClearContents
    
Dim i%, arr As Range, srr As Range

Set arr = Range("aa2:aa66653")
Set srr = Range("d2:d66653")
    
    For i = 2 To [k66365].End(3).Row
    
        If Application.WorksheetFunction.SumIfs(arr, srr, Cells(i, "d")) >= 1 Then Cells(i, "ab") = "Shortage"
        If Application.WorksheetFunction.SumIfs(arr, srr, Cells(i, "d")) = 0 Then Cells(i, "ab") = "HaveStock"
        
        Next i
        
'    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True
'    Application.StatusBar = False

End Sub


Sub PullASyardSTO() '  ADO读取其他工作簿
    Dim wb As Workbook
    Dim arr, brr(), i&, j&, k&, nrow&, crr()
    
    't = Timer
    Application.ScreenUpdating = False
    'Sheet4.Range("a1").Value = "Data collected at:" & Format(Now(), "hhmm,mm-dd-yyyy")
    
    Sheet9.Activate
    Cells.Clear
    
    Set wb = GetObject("C:\Users\jishen\Downloads\ASYARD.xlsx") '打开工作簿
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False
    ReDim brr(1 To UBound(arr), 1 To 18)
    For i = 1 To UBound(arr)
        For j = 1 To 18
            brr(i, j) = arr(i, j)
        Next
    Next
    
    With Sheet9
        .Range("A1").Resize(UBound(arr), 18) = brr
        .Columns.AutoFit
    End With
    Application.ScreenUpdating = True
    'MsgBox Format(Timer - t, "0.00" & "s")
    
End Sub





















