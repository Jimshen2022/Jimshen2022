'Pull TripDetailInformation
Sub load_TripDetails()
    
    Application.ScreenUpdating = False
    Dim i As Long
    Dim adors As New Recordset
    Sheet1.Activate
    Sheet1.Cells.Clear
    
    
    Set Db = New Connection
    Db.CursorLocation = adUseClient
    If Db.State = 1 Then Db.Close
    
    U = Sheet4.Range("a1").Value
    P = Sheet4.Range("a2").Value
    
    Db.Open "Provider =IBMDASQL.DataSource.1" & _
            ";Catalog Library List=JIMTDTA" & _
            ";Persist Security Info=True" & _
            ";Force Translate=0" & _
            ";Data Source = WVFHA" & _
            ";User ID =" & U & "" & _
            ";Password =" & P
    
    Set adors = New Recordset
    If adors.State = 1 Then adors.Close
    
     cmdtxt = "Select t2.BDITM#,t2.BDITMD,t2.BDICLS,sum(t2.BDITQT) as Qty, sum(t2.BDITCT) as Cubes, sum(t2.BDITWT) as Weight " & _
              "from (Select BDTRP#,BDITM#,BDITMD,BDICLS,BDITQT,BDITCT,BDITWT From DISTLIBQ.BTTRIPD t1 Where BDTRP# IN ('41982') order by BDTRP#,bditm#) as t2 " & _
              "group by t2.BDITM#,t2.BDITMD,t2.BDICLS "

     
    adors.Open cmdtxt, Db, 3, 3
    For i = 0 To adors.Fields.Count - 1
        Sheet1.Cells(1, i + 1) = adors.Fields(i).Name
    Next i
    
    Sheet1.Range("a2").CopyFromRecordset adors
    adors.Close
    Set adors = Nothing
    
    Sheet1.Columns("A:C").NumberFormat = "@"

    Application.ScreenUpdating = True
End Sub






