'提取所有的资料方法

Sub PullATSTO() '  ADO读取其他工作簿
    Dim wb As Workbook
    Dim arr, brr(), i&, j&, k&, nrow&, crr()
    
    't = Timer
    'Application.ScreenUpdating = False
    Sheet4.Range("a1").Value = "Data collected at:" & Format(Now(), "hhmm,mm-dd-yyyy")
    
    Sheet2.Activate
    Range("a2:af1048517").ClearContents
    
    Set wb = GetObject("C:\Users\jishen\Downloads\ASSTO.xlsx") '打开工作簿
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False
    ReDim brr(1 To UBound(arr), 1 To 18)
    For i = 1 To UBound(arr)
        For j = 1 To 18
            brr(i, j) = arr(i, j)
        Next
    Next
    Sheet2.Range("o1").Resize(UBound(arr), 18) = brr
    
    Sheet2.Activate
    Columns("o:aF").NumberFormat = "@"
    Columns("p:aF").EntireColumn.AutoFit
    Sheet2.Select
    
    'Application.ScreenUpdating = True
    'MsgBox Format(Timer - t, "0.00" & "s")
    
End Sub


'提取部分资料的方法

Sub Pull_WANEK_SN_IN_WAREHOUSE()
    Dim wb As Workbook
    Dim arr, brr(), i&, j&, k&, nrow&, crr()
    
    t = Timer
    Application.ScreenUpdating = False
    'Sheet4.Range("a1").Value = "Data collected at:" & Format(Now(), "hhmm,mm-dd-yyyy")
    
    With Sheet6
         .Cells.Clear
        
        Set wb = GetObject("C:\Users\jishen\Downloads\IW.xlsx") '打开工作簿
        nrow = wb.ActiveSheet.Range("a1048576").End(3).Row
        arr = wb.ActiveSheet.Range("a1:o" & nrow)
        wb.Close False
        
        ReDim brr(1 To UBound(arr), 1 To 4)
        For i = 1 To UBound(arr)
            
            brr(i, 1) = arr(i, 1)
            brr(i, 2) = arr(i, 2)
            brr(i, 3) = arr(i, 3)
            brr(i, 4) = arr(i, 8)
            
        Next
         .Columns("a:d").NumberFormat = "@"
         .Range("a1").Resize(UBound(arr), 4) = brr
         .Columns("a:d").EntireColumn.AutoFit
    End With
    
    Application.ScreenUpdating = True
    MsgBox Format(Timer - t, "0.00" & "s")
    
End Sub




'这里有个问题，非得对象工作簿要打开才可以

Sub 查询6()
    Dim cnn As Object, rst As Object, i&, sql$
    
    Application.ScreenUpdating = False
    
    Set cnn = CreateObject("adodb.connection")
    Set rst = CreateObject("adodb.recordset")
    
    cnn.Open "Provider = Microsoft.ACE.OLEDB.12.0;Data Source=D:\data\AT_TRX.xlsm;extended properties=""excel 12.0;HDR=YES;IMEX=1"""
    
    sql = "select * " &  _
            "from [Sheet1$] " &  _
            "where [Transaction Code] in ('152','202') and [From Location ID] like ""VR%"""
    
    
    rst.Open sql, cnn, 1, 3
    With Worksheets("LP_DATA")
         .UsedRange.ClearContents
        For i = 0 To rst.Fields.Count - 1 '输出标题
             .Cells(1, i + 1) = rst.Fields(i).Name
        Next
         .Range("a2").CopyFromRecordset rst '输出数据
    End With
    rst.Close
    cnn.Close
    Set rst = Nothing
    Set cnn = Nothing
    Application.ScreenUpdating = True
End Sub


Sub ado_HJ_trx_load()
    Dim cnn As Object, rst As Object, i&, sql$
    Application.ScreenUpdating = False
    Set cnn = CreateObject("adodb.connection")
    Set rst = CreateObject("adodb.recordset")
    cnn.Open "Provider = Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\jishen\Downloads\AT_TRX.xlsm;extended properties=""excel 12.0;HDR=YES;IMEX=1"""
    sql = "select * " &  _
            "from [Sheet1$] " &  _
            "where [Transaction Code] in ('152','202','252','254','262','304','312','321','364','372','856') and [From Location ID] like ""VR%"""
    rst.Open sql, cnn, 1, 3
    With Worksheets("LP_DATA")
         .UsedRange.ClearContents
        For i = 0 To rst.Fields.Count - 1
             .Cells(1, i + 1) = rst.Fields(i).Name
        Next
         .Range("a2").CopyFromRecordset rst
         .Columns("a:ab").EntireColumn.AutoFit
    End With
    rst.Close
    cnn.Close
    Set rst = Nothing
    Set cnn = Nothing
    Application.ScreenUpdating = True
End Sub