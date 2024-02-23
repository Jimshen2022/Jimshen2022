Sub Onhand()

Call InCharge

'On Error Resume Next
Application.ScreenUpdating = False
Dim i As Integer, j As Integer, n As Integer, m As Integer
Dim cmdtxt As String
Dim adors As New Recordset
Dim sh As Worksheet
Dim adoCN As Object
Dim strSQL As String
Dim objPivotCache As Object
Dim QDate1 As String
Dim QDate2 As String
Dim mrp As String
Dim tras As String

Sheets("Onhand").Select
If Worksheets("Onhand").Range("A1") <> "" Then
   Worksheets("Onhand").Range("A1").AutoFilter Field:=1
   Worksheets("Onhand").Range("A1").AutoFilter
End If

Range("A2:H20000").Select
Selection.ClearContents

Call MakeString

etd = Worksheets("SETTING").Range("a1").Text
sku = Worksheets("SETTING").Range("b1").Text
itcls = Worksheets("SETTING").Range("c1").Text
whs = Worksheets("SETTING").Range("d1").Text

     Set Db = New Connection
     Db.CursorLocation = adUseClient
         
     If Db.State = 1 Then Db.Close

Db.Open "Provider =IBMDASQL.DataSource.1" & _
        ";Catalog Library List=JDETSTDTA" & _
        ";Persist Security Info=True" & _
        ";Force Translate=0" & _
        ";Data Source = MILPROD" & _
        ";User ID =LLSEW1" & _
        ";Password =LLSEW1"

     Worksheets("Onhand").Select
     Set adors = New Recordset
     If adors.State = 1 Then adors.Close
     
     cmdtxt = "SELECT SLQNTY.HOUSE, SLQNTY.ITNBR as ITEM, ITMRVAL0.ITDSAD, " & _
              "ITMRVAL0.ITCLAD, " & _
              "SLQNTY.LQNTY, SLQNTY.LLOCN, ITMRVAL0.ITYPAD " & _
              "FROM AMFLIBL.SLQNTY SLQNTY, AMFLIBL.ITMRVAL0 ITMRVAL0 " & _
              "WHERE ITMRVAL0.ITNOAD = SLQNTY.ITNBR AND ITMRVAL0.STIDAD = SLQNTY.HOUSE " & _
              "AND SLQNTY.LLOCN in ('FA00') AND SLQNTY.ITNBR like '%UN%' AND ITMRVAL0.ITCLAD not in ('ZKIZ') "
        
        
        If whs <> "" Then
              cmdtxt = cmdtxt & " AND SLQNTY.HOUSE in " & whs & ""
        End If
                
        If sku <> "" Then
              cmdtxt = cmdtxt & " AND SLQNTY.ITNBR in " & sku & ""
        End If
        
        If itcls <> "" Then
              cmdtxt = cmdtxt & " AND ITMRVAL0.ITCLAD in " & itcls & ""
        End If
     
     adors.Open cmdtxt, Db, 3, 3
     
     For i = 0 To adors.Fields.Count - 1
         Worksheets("Onhand").Cells(1, i + 1) = adors.Fields(i).Name
     Next i
      
    Worksheets("Onhand").Range("A2").CopyFromRecordset adors

Call Load_FinalMO
Worksheets("Setting").Select
Application.ScreenUpdating = True
'MsgBox "Onhand Downloaded!"

End Sub

