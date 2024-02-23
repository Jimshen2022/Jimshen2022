'D:\Document\01-Wanvog\10-CC\Anual CC\2021\HJ vs Mapics Comparing - 2021.xlsx

'Pull SN IN WAREHOUSE
Sub Pull_SN_IN_WAREHOUSE2()                                    

    Dim wb As Workbook, arr
    Application.ScreenUpdating = False
    
    Sheet2.Activate
    Sheet2.Cells.Clear
    Set wb = GetObject("C:\Users\jishen\Downloads\AS_INWAREHOUSE.xlsx") '打开工作簿
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False

    Columns("a:p").NumberFormat = "@"
    Sheet2.Range("a1").Resize(UBound(arr), UBound(arr, 2)) = arr
    Columns("a:p").EntireColumn.AutoFit
    Erase arr
    Application.ScreenUpdating = True
    
End Sub




'范例46-5 使用SQL连接取得数据  --- '159669 rows 15.59s
Sub UsingSQL()
    Dim SQL As String
    Dim j As Integer
    Dim r As Integer
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
    With Sheet1
         .Cells.clear
        Set cnn = New ADODB.Connection
        With cnn
             .Provider = "Microsoft.ACE.OLEDB.12.0"
             .Connectionstring = "Extended Properties = Excel 12.0;" &  _
                    "Data Source = " & Thisworkbook.Path & "\数据.xlsx"
             .Open
        End With
        
        Set rs = New ADODB.Recordset
        SQL = "SELECT * FROM [Sheet1$]"
        rs.Open SQL, cnn, adOpenKeyset, adLockOptimistic
        
        For j = 0 To rs.Fields.Count - 1
             .Cells(1, j + 1) = rs.Fields(j).Name
        Next
        
        r =  .cells(.rows.count, 1).End(xlUp).row
         .Range("a" & r + 1).CopyFromRecordset rs
    End With
    rs.close
    cnn.close
    Set rs = Nothing
    Set cnn = Nothing
    
End Sub


'Pull SN_LOADED
Sub Pull_SN_LOADED2()
    On Error Resume Next
    Dim wb As Workbook
    Dim arr, brr(), i&, j&, k&, m&, nrow&, crr()
    Application.ScreenUpdating = False
    Sheet2.Activate
    Set wb = GetObject("C:\Users\jishen\Downloads\AS_LOADED.xlsx") '打开工作簿
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False
    ReDim brr(1 To UBound(arr) - 1, 1 To UBound(arr, 2))
    For i = 2 To UBound(arr)
        If arr(i, 1) <> "" Then
            m = m + 1
            For j = 1 To 16
                brr(m, j) = arr(i, j)
            Next
        End If
    Next
    Columns("a:p").NumberFormat = "@"
    Sheet2.Range("a1048576").End(3).Offset(1, 0).Resize(UBound(arr) - 1, UBound(arr, 2)) = brr
    Application.ScreenUpdating = True
   
End Sub

'Pull SN HOLD
Sub Pull_SN_HOLD2()

    Dim wb As Workbook
    Dim arr, brr(), i&, j&, k&, m&, nrow&, crr()
    
    Application.ScreenUpdating = False
    Sheet2.Activate
    
    Set wb = GetObject("C:\Users\jishen\Downloads\AS_HOLD.xlsx") '打开工作簿
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False
    ReDim brr(1 To UBound(arr) - 1, 1 To UBound(arr, 2) + 1)
    For i = 2 To UBound(arr)
        
        If arr(i, 9) <> "QA001VD1" And arr(i, 5) <> "Orphaned" Then
            m = m + 1
            For j = 1 To 16
                brr(m, j) = arr(i, j)
            Next
        End If
    Next
    Columns("a:p").NumberFormat = "@"
    Sheet2.Range("a1048576").End(3).Offset(1, 0).Resize(UBound(arr) - 1, UBound(arr, 2)) = brr
    Columns("a:p").EntireColumn.AutoFit
    
    Application.ScreenUpdating = True
    'MsgBox Format(Timer - t, "0.00" & "s")
    
End Sub


'Pull SN ORPHANED

Sub Pull_SN_ORPHANED2()
    Application.ScreenUpdating = False
    
    Dim wb As Workbook
    Dim arr, brr(), i&, j&, k&, m&, nrow&, crr()
    Sheet7.Activate
    Cells.Clear
        
    Set wb = GetObject("C:\Users\jishen\Downloads\AS_ORPHANED.xlsx") '打开工作簿
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False
       
    ReDim brr(1 To UBound(arr), 1 To 6)
        For i = 1 To UBound(arr)
            brr(i, 1) = arr(i, 2)
            brr(i, 2) = arr(i, 3)
            brr(i, 3) = arr(i, 5)
            brr(i, 4) = arr(i, 9)
            brr(i, 5) = arr(i, 11)
            brr(i, 6) = arr(i, 8)
        Next
        
        With Sheet7
         .Columns("a:f").NumberFormat = "@"
         .Range("a1").Resize(UBound(arr), 6) = brr
         .Columns("a:f").EntireColumn.AutoFit
        End With
    Sheet7.Range("a1:f" & Range("a65563").End(xlUp).Row).Sort [b1], xlAscending, [a1], xlDscending, , , , xlYes
    Application.ScreenUpdating = True

End Sub

'Pull SN IN WAREHOUSE
Sub Pull_ASSTO()
    Application.ScreenUpdating = False
    
    Dim wb As Workbook
    Dim arr, brr(), i&, j&, k&, m&, nrow&, crr()
    Sheet4.Activate
    Cells.Clear
  
    Set wb = GetObject("C:\Users\jishen\Downloads\ASSTO.xlsx") '打开工作簿
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False

    Columns("a:r").NumberFormat = "@"
    Sheet4.Range("a1").Resize(UBound(arr), UBound(arr, 2)) = arr
    Columns("a:r").EntireColumn.AutoFit
   
    Application.ScreenUpdating = True
    
End Sub
'Pull STO<>SNA
Sub Pull_SNA_NOT_BALANCE()
    On Error Resume Next
    Dim wb As Workbook, arr
    Application.ScreenUpdating = False
    
    Sheet1.Activate
    Sheet1.Cells.Clear
    Set wb = GetObject("C:\Users\jishen\Downloads\Search_STO_and_SNA_Balance.xlsx") '打开工作簿
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False

    Columns("a:p").NumberFormat = "@"
    Sheet1.Range("a1").Resize(UBound(arr), UBound(arr, 2)) = arr
    Columns("a:p").EntireColumn.AutoFit
    Erase arr
    Application.ScreenUpdating = True
    
End Sub

Sub unfilter2()
    '取消筛选
    Application.ScreenUpdating = False
    On Error Resume Next
    Dim i%, sht As Worksheet
    
    For Each sht In Worksheets
        ' 如果当前工作表为筛选模式，则取消
        If sht.AutoFilterMode = True Then sht.AutoFilterMode = 0
        ' 如果当前工作表没有筛选，则加上筛选
        If sht.AutoFilterMode = False Then sht.Range("a1").AutoFilter Field:=1
    Next
     Application.ScreenUpdating = True
        
End Sub

'Sub DistinctSN2()
'
'    Application.ScreenUpdating = False
'    '    Application.Calculation = xlCalculationManual
'    '    Application.StatusBar = "Calculating, please wait ......"
'
'    Sheet3.Activate
'    Sheet3.Cells.Clear
'
'    Set cnn = CreateObject("adodb.connection")
'    cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=""Excel 12.0;HDR=YES"";Data Source=" & ThisWorkbook.FullName
'    Sql = "SELECT Distinct [Serial Number],[Warehouse],[Item Number],[Location] " & _
'          "FROM [ALL_SN$] where [Location] NOT LIKE 'RP04%' "
'
'
'    Sheet3.Range("a2").CopyFromRecordset cnn.Execute(Sql)
'    cnn.Close
'    Set cnn = Nothing
'    'Columns("t:t").NumberFormat = "m/d/yyyy"
'
'    '    Application.Calculation = xlCalculationAutomatic
'   Application.ScreenUpdating = True
'    '    Application.StatusBar = False
'
'End Sub


Sub DistinctSN1()
    
    Application.ScreenUpdating = False
    't = Timer
    Dim aKey, aItem, kRes, iRes, arr, i&, j&, d As Object

    Sheet3.Activate
    Sheet3.Cells.Clear
    arr = Sheet2.Range("a2:p" & Rows.Count).CurrentRegion
    
    Set d = CreateObject("scripting.dictionary")
    d.CompareMode = vbTextCompare '不区分字母大小写
    
    For i = 2 To UBound(arr) '遍历数组arr
        If Not arr(i, 9) Like "RP0%" Then
            d(arr(i, 2)) = arr(i, 1) & ";" & arr(i, 3) & ";" & arr(i, 5) & ";" & arr(i, 6) & ";" & arr(i, 8) & ";" & arr(i, 9) & ";" & arr(i, 11) '将SN 作为key，装入字典
        End If
    Next
  
    '遍历字典的items,存放于aRes数组
    aItem = d.items
    aKey = d.Keys
    
    ReDim iRes(1 To d.Count + 1, 1 To 7)  '结果数组
    ReDim kRes(1 To d.Count + 1, 1 To 1)  '结果数组
    
    
    For i = 0 To UBound(aKey)
        For j = 1 To 1
           kRes(i + 1, j) = aKey(i)
        Next
    Next
       
    With Sheet3
        .Columns("a:g").NumberFormat = "@"
        .Columns("a:h").AutoFit
        .Range("a2").Resize(UBound(iRes), 1) = kRes
    End With
    
    
    For i = 0 To UBound(aItem)
        For j = 0 To 6
           iRes(i + 1, j + 1) = Split(aItem(i), ";")(j)
        Next
    Next
    
    With Sheet3
        .Columns("a:g").NumberFormat = "@"
        .Columns("a:h").AutoFit
        .Range("b2").Resize(UBound(iRes), 7) = iRes
        .Range("a1:h1").Value = Array("Serial Number", "'Warehouse", "Item Number", "Active Status", "Master Status", "Master MO/PO", "Location", "Received Date")
    End With
    
    Set d = Nothing
    Erase aItem
    Erase aKey
    Erase kRes
    Erase iRes
    Erase arr
    'ThisWorkbook.Save
    Application.ScreenUpdating = True
    'MsgBox Format(Timer - t, "0.00" & "s")
    

End Sub





































