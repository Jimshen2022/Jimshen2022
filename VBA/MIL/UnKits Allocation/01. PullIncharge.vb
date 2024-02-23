Sub InCharge()

Application.Calculation = xlManual
'On Error Resume Next
Application.ScreenUpdating = False

Dim i As Integer, j As Integer, n As Integer, m As Integer
Dim cmdtxt As String
Dim adors As New Recordset
Dim sh As Worksheet
Dim adoCN As Object
Dim strSQL As String
Dim objPivotCache As Object

Sheets("MatInChar").Select
Range("A3:M60000").Select
Selection.ClearContents

     Worksheets("SETTING").Select
     
     Call MakeString
     x = Worksheets("SETTING").Range("B1").Text
     y = Worksheets("SETTING").Range("A1").Text
     Item = Worksheets("SETTING").Range("C1").Text
     itcls = Worksheets("SETTING").Range("D1").Text
     
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
     
     Worksheets("MatInChar").Select
     Set adors = New Recordset
     If adors.State = 1 Then adors.Close
     
     cmdtxt = "SELECT B.BXDSTID, B.BXDPARENTITEMCLASS as P_ITCLS, B.BXDPARENTITEMNUMBER as P_ITEM," & _
                "B.BXDCOMPONENTITEMCLASS as T_ITCLS, B.BXDCOMPONENTITEMNUMBER as T_ITEM, " & _
                "B.BXDCOMPONENTITEMDESCRIPTION as T_DES, B.BXDITEMTYPECODE as C_Type," & _
                "B.BXDQUANTITYPERUNIT as C_CONS, B.BXDUNITOFMEASURE as C_UNIT " & _
                "FROM RGNFILL.TBL_BOM_EXTRACT_DETAIL B " & _
                "WHERE B.BXDQUANTITYPERUNIT <>0 " & _
                "AND B.BXDPARENTITEMNUMBER like '%UN%' AND B.BXDPARENTITEMCLASS not in ('ZKIS','ZKIZ') " & _
                "AND B.BXDCOMPONENTITEMCLASS in ('RCT') AND B.BXDPARENTITEMNUMBER not like '%PACK%'"
        
     cmdtxt = cmdtxt
     adors.Open cmdtxt, Db, 3, 3
     
     For i = 0 To adors.Fields.Count - 1
         Worksheets("MatInChar").Cells(2, i + 1) = adors.Fields(i).Name
     Next i

     Worksheets("MatInChar").Range("A3").CopyFromRecordset adors
       
Set ws = Worksheets("MatInChar")
LR = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
Calculate

Worksheets("MatInChar").Select
Application.ScreenUpdating = True
'MsgBox "MatInChar Downloaded!
'Application.Calculation = xlAutomatic
End Sub

