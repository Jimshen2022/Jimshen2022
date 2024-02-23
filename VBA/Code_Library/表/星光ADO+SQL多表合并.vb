Sub bySQL()
    Dim cnADO As Object
    Dim rsADO As Object
    Dim strSQL As String
    Dim i As Long, strShtName, aShtName
    Set cnADO = CreateObject("ADODB.Connection")
    Set rsADO = CreateObject("ADODB.Recordset")
    cnADO.Open "Provider=Microsoft.ACE.OLEDB.12.0;" _
             & "Extended Properties=Excel 12.0;" _
             & "Data Source=" & ThisWorkbook.FullName
    aShtName = Split("一部门,二部门,三部门,四部门,后勤部", ",")
    For Each strShtName In aShtName '多表合并语句
        strSQL = strSQL & "SELECT 姓名,工号 ,'" & strShtName & " ' AS 工作表名称 FROM [" & strShtName & "$]  UNION ALL "
    Next
    Set rsADO = cnADO.Execute(Left(strSQL, Len(strSQL) - 10))
    Cells.ClearContents
    For i = 0 To rsADO.Fields.Count - 1
        Cells(1, i + 1) = rsADO.Fields(i).Name
    Next i
    Range("a2").CopyFromRecordset rsADO
    rsADO.Close
    cnADO.Close
    Set cnADO = Nothing
    Set rsADO = Nothing
End Sub