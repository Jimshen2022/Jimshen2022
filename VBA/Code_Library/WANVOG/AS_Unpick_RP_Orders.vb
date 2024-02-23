
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