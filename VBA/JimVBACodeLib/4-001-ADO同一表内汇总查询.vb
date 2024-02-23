Sub MIL_OH_PULLOUT_AND_MAKE_CC_LIST()
    
    Application.ScreenUpdating = False
    t = Timer
    Call cPullOHList
    ThisWorkbook.Save
    
    Dim i&, cnn As Object, rs As Object, sql$, ccDate As Date
    Sheet3.Activate
    Sheet3.Range("a2:l1048576").Cells.Clear
    ccDate = Sheet4.Range("b8").Value
    
    Set cnn = CreateObject("adodb.connection")
    Set rs = CreateObject("ADODB.Recordset")
    
    cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=""Excel 12.0;HDR=YES"";Data Source=" & ThisWorkbook.FullName
    
            
        sql = "select t1.[CC Date],t1.[WH],t1.[Locations],t2.[ITNBR],t2.[LQNTY],t2.[ORDNO],t2.[LBHNO]" & _
            "from [CC Schedule$] as t1 left join (select * from [OnHand$] as x) t2  " & _
            "on t2.[LLOCN]=t1.[Locations] and t2.[HOUSE]=t1.[WH] " & _
            "where t1.[CC Date] = #" & ccDate & "#"
    
    Set rs = cnn.Execute(sql)
'    For i = 0 To rs.Fields.Count - 1
'        Cells(1, i + 1) = rs.Fields(i).Name
'    Next i
'            "where t1.[LLOCN]=t2.[Locations] and t2.[CC Date] in (" & ccDate & ")"
    Sheet3.Columns("a:d").NumberFormat = "@"
'    Sheet3.Columns("e:e").NumberFormat = "#,###"
    Sheet3.Columns("f:g").NumberFormat = "@"
    Sheet3.Range("a2").CopyFromRecordset cnn.Execute(sql)
    
    cnn.Close
    
    Set cnn = Nothing
    Set rs = Nothing
    
        nrow = Sheet3.Range("c1048576").End(3).Row
    'format
    With Sheet3
        .Range("e1:e" & nrow).Font.ColorIndex = 3
        .Range("h1:h" & nrow).Font.ColorIndex = 5
        .Range("a1:k" & nrow).Borders.LineStyle = xlContinuous
        .Range("a1:k" & nrow).HorizontalAlignment = xlCenter
        .Range("i2").Formula = "=h2-e2"
        .Range("j2").Formula = "=abs(i2)"
        .Range("k2").Formula = "=if(i2=0,""NoVariance"",if(i2<0,""Inv.Loss"",""Inv.Gain""))"
        .Range("i2:k2").AutoFill Destination:=Range("i2:k" & nrow)
        .Range("i1:k" & nrow).Interior.ColorIndex = 15
        .Range("i1:k1").Value = Array("Variance", "abs(Variance)", "Type")
    End With
    
    Sheet3.Columns("a:m").AutoFit
    Sheet3.Range("a1:k1").AutoFilter
    ThisWorkbook.Save
    
    Application.ScreenUpdating = True
    MsgBox "Updated Successful~    " & Format(Timer - t, "0.00" & "s")
    
End Sub

'Sub formulasAndFormat()
'    Application.ScreenUpdating = False
'
'    nrow = Sheet3.Range("c1048576").End(3).Row
'    'format
'    With Sheet3
'        .Range("e1:e" & nrow).Font.ColorIndex = 3
'        .Range("h1:h" & nrow).Font.ColorIndex = 5
'        .Range("a1:k" & nrow).Borders.LineStyle = xlContinuous
'        .Range("a1:k" & nrow).HorizontalAlignment = xlCenter
'        .Range("i2").Formula = "=h2-e2"
'        .Range("j2").Formula = "=abs(i2)"
'        .Range("k2").Formula = "=if(i2=0,""NoVariance"",if(i2<0,""Inv.Loss"",""Inv.Gain""))"
'        .Range("i2:k2").AutoFill Destination:=Range("i2:k" & nrow)
'        .Range("i1:k" & nrow).Interior.ColorIndex = 15
'
'    End With
'
'    Application.ScreenUpdating = True
'End Sub

