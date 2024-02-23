Sub putaway_time_666()
    Dim i&, nrow&, h!, k!
    
    With Sheet10
        nrow =  .[a1048576].End(3).Row
        Sheet6.Range("q2").ClearContents
         .Range("u2:u" & nrow) =  .Range("u2:u" & nrow).Value2
        
        h = Application.WorksheetFunction.Max(.Range("u2:u" & nrow))
        k = Application.WorksheetFunction.Min(.Range("u2:u" & nrow))
        Sheet6.Range("q2").Value = (h - k) * 24
    End With
    
End Sub

