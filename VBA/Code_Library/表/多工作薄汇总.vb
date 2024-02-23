Sub ADO法()
    Dim cnn As Object, MyPath$, MyFile$, SQL$, m%
    MyPath = ThisWorkbook.Path & ""
    MyFile = Dir(MyPath & "*.xlsx")
    Application.ScreenUpdating = False
    Cells.ClearContents
    [a1:d1] = Array("序号", "村民小组", "户主姓名", "家庭人口 数")
    Set cnn = CreateObject("adodb.connection")
    Do While Len(MyFile)
        m = m + 1
        If m = 1 Then
            cnn.Open "provider=microsoft.ace.oledb.12.0;extended properties=excel 12.0;data source=" & MyPath & MyFile
            SQL = "select * from [Sheet0$a3:d] where 村民小组 is not null"
        Else
            SQL = "select * from [Excel 12.0;Database=" & MyPath & MyFile & "].[Sheet0$a3:d] where 村民小组 is not null"
        End If
        Range("a" & Rows.Count).End(xlUp).Offset(1).CopyFromRecordset cnn.Execute(SQL)
        MyFile = Dir
    Loop
    With [a1].CurrentRegion
         .Value =  .Value
    End With
    cnn.Close
    Set cnn = Nothing
    Application.ScreenUpdating = True
End Sub

'ADO方法2

Sub Wanek1()
    Dim cnn As Object, rst As Object, i&, sql$
    
    Application.ScreenUpdating = False
    
    Set cnn = CreateObject("adodb.connection")
    Set rst = CreateObject("adodb.recordset")
    
    cnn.Open "Provider = Microsoft.ACE.OLEDB.12.0;Data Source=\\10.141.100.133\VNData\UPH FG Warehouse\Public\Luanna\Container_transfer\WNK1.xlsx;extended properties=""excel 12.0;HDR=YES;IMEX=1"""
    
    sql = "select * " &  _
            "from [Sheet1$a2:q] "
    
    
    rst.Open sql, cnn, 1, 3
    With Worksheets("Summary")
         .Range("a3:q10000").ClearContents
        '        For i = 0 To rst.Fields.Count - 1 '输出标题
        '             .Cells(1, i + 1) = rst.Fields(i).Name
        '        Next
         .Range("a3").CopyFromRecordset rst '输出数据
    End With
    rst.Close
    cnn.Close
    Set rst = Nothing
    Set cnn = Nothing
    Call Wanek2
    Call Wanek3
    ThisWorkbook.Save
    MsgBox "updated"
    Application.ScreenUpdating = True
End Sub

Sub Wanek2()
    Dim cnn As Object, rst As Object, i&, sql$, nrow&
    
    Application.ScreenUpdating = False
    nrow = Sheet5.Range("a1048576").End(3).Row
    Set cnn = CreateObject("adodb.connection")
    Set rst = CreateObject("adodb.recordset")
    
    cnn.Open "Provider = Microsoft.ACE.OLEDB.12.0;Data Source=\\10.141.100.133\VNData\UPH FG Warehouse\Public\Luanna\Container_transfer\WNK2.xlsx;extended properties=""excel 12.0;HDR=YES;IMEX=1"""
    
    sql = "select * " &  _
            "from [Sheet1$a2:q] "
    
    
    rst.Open sql, cnn, 1, 3
    With Worksheets("Summary")
        '.Range("a3:q10000").ClearContents
        '        For i = 0 To rst.Fields.Count - 1 '输出标题
        '             .Cells(1, i + 1) = rst.Fields(i).Name
        '        Next
         .Range("a1048576").End(3).Offset(1).CopyFromRecordset rst '输出数据
    End With
    rst.Close
    cnn.Close
    Set rst = Nothing
    Set cnn = Nothing
    Application.ScreenUpdating = True
End Sub

Sub Wanek3()
    Dim cnn As Object, rst As Object, i&, sql$, nrow&
    
    Application.ScreenUpdating = False
    nrow = Sheet5.Range("a1048576").End(3).Row
    Set cnn = CreateObject("adodb.connection")
    Set rst = CreateObject("adodb.recordset")
    
    cnn.Open "Provider = Microsoft.ACE.OLEDB.12.0;Data Source=\\10.141.100.133\VNData\UPH FG Warehouse\Public\Luanna\Container_transfer\WNK3.xlsx;extended properties=""excel 12.0;HDR=YES;IMEX=1"""
    
    sql = "select * " &  _
            "from [Sheet1$a2:q] "
    
    
    rst.Open sql, cnn, 1, 3
    With Worksheets("Summary")
        '.Range("a3:q10000").ClearContents
        '        For i = 0 To rst.Fields.Count - 1 '输出标题
        '             .Cells(1, i + 1) = rst.Fields(i).Name
        '        Next
         .Range("a1048576").End(3).Offset(1).CopyFromRecordset rst '输出数据
    End With
    rst.Close
    cnn.Close
    Set rst = Nothing
    Set cnn = Nothing
    Application.ScreenUpdating = True
End Sub


