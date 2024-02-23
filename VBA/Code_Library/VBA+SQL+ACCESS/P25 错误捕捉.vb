Sub 错误捕捉()
    
    Dim myData As String
    Dim myTable As String
    Dim cnn As New ADODB.Connection
    myData = ThisWorkbook.Path & "\成绩管理.accdb"
    myTable = "期末成绩1"
    
    With cnn
         .Provider = "microsoft.ace.oledb.12.0"
         .Open myData
    End With
    
    On Error Resume Next '遇到错误，继续向下执行
    cnn.Execute "drop table" & myTable
    
    If Err.Number <> 0 Then '出错了，表示表不存在
        MsgBox Err.Description
    Else
        MsgBox "该表存在."
    End If
    
    cnn.Close
    Set cnn = Nothing
    
End Sub
