Sub NoSiteInColumnASorting() 'update the site sheet
    
    Sheet8.Activate
    Set CNN = CreateObject("adodb.connection")
    CNN.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=Excel 12.0;Data Source=" & ThisWorkbook.FullName
    Sql = "select Distinct [WH+EMP+SUP] " _
             & "from [Wanek3_DC_OUT$] " _
             & "where [SITE]=""VIEW"" " _
             & "Order by [WH+EMP+SUP]"
    
    Sheet8.[a33653].End(3).Offset(1, 0).CopyFromRecordset CNN.Execute(Sql)
    CNN.Close
    Set CNN = Nothing
    Sheet8.Select
    MsgBox "updated"
    
End Sub




Sub STONoSiteInColumnASorting() 'update STO A column site sheet
    
    Sheet3.Activate
    Set CNN = CreateObject("adodb.connection")
    CNN.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=Excel 12.0;Data Source=" & ThisWorkbook.FullName
    Sql = "select Distinct [Site], [Location Id] " _
             & "from [STO$] " _
             & "where [SITE]=""VIEW"" " _
             & "Order by [Site]"
    
    Sheet11.[a33653].End(3).Offset(1, 0).CopyFromRecordset CNN.Execute(Sql)
    CNN.Close
    Set CNN = Nothing
    
    Sheet11.Select
    MsgBox "updated"
    
End Sub