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
