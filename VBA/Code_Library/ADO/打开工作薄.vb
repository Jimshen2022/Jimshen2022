


Sub InTransitlist()    
    
    Application.ScreenUpdating = False
    Dim i&, cnn As Object, rs As Object, sql$
    Sheet12.Activate
    Sheet12.Cells.Clear
    
    Set cnn = CreateObject("adodb.connection")
    Set rs = CreateObject("ADODB.Recordset")
    
    cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=""Excel 12.0;HDR=YES"";Data Source=" & ThisWorkbook.FullName
    
    sql = "select [Type],[Joint] as [WH_CNT],[Item Number],Sum([Qty]) as [Q'ty]" &  _
            "from [SN$] " &  _
            "where [Type] = ""InTransit""  " &  _
            "group by [Type],[Joint],[Item Number]"
    
    Set rs = cnn.Execute(sql)
    For i = 0 To rs.Fields.Count - 1
        Cells(1, i + 1) = rs.Fields(i).Name
    Next i
    
    Sheet12.Range("a2").CopyFromRecordset cnn.Execute(sql)
    
    cnn.Close
    
    Set cnn = Nothing
    Set rs = Nothing
    Sheet12.Columns("a:d").AutoFit
    Sheet12.Range("a1:d1").AutoFilter
    
    Application.ScreenUpdating = True
    
End Sub



Sub loading_unpick_orders()
    
    '    Application.ScreenUpdating = False
    '    Application.Calculation = xlCalculationManual
    '    Application.StatusBar = "Calculating, please wait ......"
    
    Sheet7.Activate
    Range("a2:az66653").ClearContents
    
    
    Set cnn = CreateObject("adodb.connection")
    cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=""Excel 12.0;HDR=YES"";Data Source=" & ThisWorkbook.FullName
    sql = "select * from [Trx$] where [Transaction Code] in ('152','202') and [From Location ID] like ""VR%""  "
    
    
    
    Sheet7.Range("a2").CopyFromRecordset cnn.Execute(sql)
    cnn.Close
    Set cnn = Nothing
    'Columns("t:t").NumberFormat = "m/d/yyyy"
    
    '    Application.Calculation = xlCalculationAutomatic
    '    Application.ScreenUpdating = True
    '    Application.StatusBar = False
    
End Sub