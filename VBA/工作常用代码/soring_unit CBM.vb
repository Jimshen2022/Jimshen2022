Sub wvgUniteCBMsorting()    
    'Check STO sheet unit cbm 为999999的，999999代表没有找到对应item的unit CBM
    
    Sheet2.Activate
    Set CNN = CreateObject("adodb.connection")
    CNN.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=Excel 12.0;Data Source=" & ThisWorkbook.FullName
    Sql = "select Distinct [Item Number] " _
             & "from [1.STO$] " _
             & "where [Unit CBM] = 999999  " _
             & "Order by [Item Number]"
    
    Sheet12.[a1048576].End(3).Offset(1, 0).CopyFromRecordset CNN.Execute(Sql)
    CNN.Close
    Set CNN = Nothing
    Sheet12.Select
    'MsgBox "updated"
    
End Sub