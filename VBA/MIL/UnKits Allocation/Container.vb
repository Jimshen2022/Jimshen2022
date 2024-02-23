Sub Container()
'On Error Resume Next
Application.ScreenUpdating = False
Dim i As Integer, j As Integer, n As Integer, m As Integer
Dim cmdtxt As String
Dim adors As New Recordset
Dim sh As Worksheet
Dim adoCN As Object
Dim strSQL As String
Dim objPivotCache As Object

If Worksheets("Container").Range("A1") <> "" Then
Worksheets("Container").Range("A1").AutoFilter Field:=1
Worksheets("Container").Range("A1").AutoFilter
End If
    
Sheets("Container").Select
Columns("A:F").Select
Selection.ClearContents

Application.Calculation = xlManual

     Set Db = New Connection
         Db.CursorLocation = adUseClient
         
      If Db.State = 1 Then Db.Close
        
Db.Open "Provider =IBMDASQL.DataSource.1" & _
        ";Catalog Library List=JDETSTDTA" & _
        ";Persist Security Info=True" & _
        ";Force Translate=0" & _
        ";Data Source = MILPROD" & _
        ";User ID =LLSEW1 " & _
        ";Password =LLSEW1 "

         
     Worksheets("Container").Select
     Set adors = New Recordset
     If adors.State = 1 Then adors.Close
     
  
      cmdtxt = "SELECT a.WCHCONTAINERNUMBER, a.WCHORIGIN, a.WCHDESTINATION, a.WCHCONTAINERSTATUS, a.WCHLASTMAINTENANCETIMESTAMP " & _
                "FROM LLUSAF.TBL_WVCONTAINER_HDR a " & _
                "WHERE (a.WCHCONTAINERSTATUS='P') AND (a.WCHCONTAINERNUMBER Not Like 'AIR%') AND (a.WCHDESTINATION Not In ('21','CNW','C')) "
                    
     adors.Open cmdtxt, Db, 3, 3
     
     For i = 0 To adors.Fields.Count - 1
         Worksheets("Container").Cells(1, i + 1) = adors.Fields(i).Name
     Next i
      
Worksheets("Container").Range("A2").CopyFromRecordset adors
    
Set ws = Sheets("Container")
Set pr = Sheets("Setting")
LR = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
pr.Range("F6:F10000").ClearContents      'clear old data

j = 6
For i = 2 To LR
pr.Cells(j, 6) = ws.Range("A" & i)
j = j + 1
Next i

Call DetailsContainer
'Application.Calculation = xlManual
       
Worksheets("Setting").Select
Application.ScreenUpdating = True
MsgBox "Container Downloaded!"

End Sub


'--------------------------------------------------------------------------------------------------------------------------------

Sub DetailsContainer()

'On Error Resume Next
Application.ScreenUpdating = False
Dim i As Integer, j As Integer, n As Integer, m As Integer
Dim cmdtxt As String
Dim adors As New Recordset
Dim sh As Worksheet
Dim adoCN As Object
Dim strSQL As String
Dim objPivotCache As Object

Sheets("Container").Select
Columns("H:P").Select
    Selection.ClearContents
Call MakeString
Worksheets("SETTING").Select
cont = Worksheets("SETTING").Range("F1").Text
Application.Calculation = xlManual

     Set Db = New Connection
         Db.CursorLocation = adUseClient
         
      If Db.State = 1 Then Db.Close
        
Db.Open "Provider =IBMDASQL.DataSource.1" & _
        ";Catalog Library List=JDETSTDTA" & _
        ";Persist Security Info=True" & _
        ";Force Translate=0" & _
        ";Data Source = MILPROD" & _
        ";User ID =LLSEW1 " & _
        ";Password =LLSEW1"
         
     Worksheets("Container").Select
     Set adors = New Recordset
     If adors.State = 1 Then adors.Close
     
  
      cmdtxt = "SELECT b.WCIORIGIN as Warehouse,b.WCIORDER as Order,b.WCICONTAINERNUMBER as container ,a.ITNOIM as Item, a.B2AAS3 as Weight, a.B2CQCD as Unit, a.B2Z95S as Cube, a.B2Z93R as Unit1,b.WCIQUANTITYLOADED as Quantity " & _
                "FROM AMFLIBL.ITEMASL0 a,LLUSAF.WVCNTID b  " & _
                "WHERE a.ITNOIM = b.WCIITEMNUMBER and b.WCICONTAINERNUMBER in " & cont & " AND a.ITNOIM like '%UN%' " & _
                "ORDER BY b.WCICONTAINERNUMBER"
    
       
     adors.Open cmdtxt, Db, 3, 3
     
     For i = 0 To adors.Fields.Count - 1
         Worksheets("Container").Cells(1, i + 8) = adors.Fields(i).Name
     Next i
      

Worksheets("Container").Range("H2").CopyFromRecordset adors
    
Application.ScreenUpdating = True

End Sub

