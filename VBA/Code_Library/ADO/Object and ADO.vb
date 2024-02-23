'范例46-2 使用getobject函数取得数据   '159669 rows 8.59s
Sub UseGetObject()
    Dim Wb As Workbook
    Dim Temp As String
    Temp = Thisworkbook.path & "\数据.xlsx"
    Set Wb = GetObject(Temp)
    With Wb.sheets(1).Range("a1").Currentregion
        range("a1").Resize(.rows.count,  .columns.count) =  .Value
    End With
    Wb.Close False
    Set Wb = Nothing
End Sub


'范例46-5 使用SQL连接取得数据  --- '159669 rows 15.59s
Sub UsingSQL()
    Dim SQL As String
    Dim j As Integer
    Dim r As Integer
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
    With Sheet1
         .Cells.clear
        Set cnn = New ADODB.Connection
        With cnn
             .Provider = "Microsoft.ACE.OLEDB.12.0"
             .Connectionstring = "Extended Properties = Excel 12.0;" &  _
                    "Data Source = " & Thisworkbook.Path & "\数据.xlsx"
             .Open
        End With
        
        Set rs = New ADODB.Recordset
        SQL = "SELECT * FROM [Sheet1$]"
        rs.Open SQL, cnn, adOpenKeyset, adLockOptimistic
        
        For j = 0 To rs.Fields.Count - 1
             .Cells(1, j + 1) = rs.Fields(j).Name
        Next
        
        r =  .cells(.rows.count, 1).End(xlUp).row
         .Range("a" & r + 1).CopyFromRecordset rs
    End With
    rs.close
    cnn.close
    Set rs = Nothing
    Set cnn = Nothing
    
End Sub
