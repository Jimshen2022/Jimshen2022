Sub ByADO_SQL()
    Dim cnADO As Object
    Dim rsADO As Object
    Dim strSQL As String
    Dim i As Long
    Set cnADO = CreateObject("ADODB.Connection")
    Set rsADO = CreateObject("ADODB.Recordset")
    cnADO.Open "Provider=Microsoft.ACE.OLEDB.12.0;" _
             & "Extended Properties=Excel 12.0;" _
             & "Data Source=" & ThisWorkbook.FullName
    strSQL = "SELECT *  FROM [A$] " '//此处写入SQL代码
    Set rsADO = cnADO.Execute(strSQL)
    '//将工作表名称修改为实际放置查询数据的工作表名称▼
    Worksheets("工作表名称").Select
    Cells.ClearContents
    For i = 0 To rsADO.Fields.Count - 1
        Cells(1, i + 1) = rsADO.Fields(i).Name
    Next i
    Range("A2").CopyFromRecordset rsADO
    rsADO.Close
    cnADO.Close
    Set cnADO = Nothing
    Set rsADO = Nothing
End Sub