-- Wanek OpenMO
Sub ScheduleDue()
'On Error Resume Next
Application.ScreenUpdating = False
Dim i As Integer, j As Integer, n As Integer, m As Integer
Dim cmdtxt As String
Dim adors As New Recordset
Dim sh As Worksheet
Dim adoCN As Object
Dim strSQL As String
Dim objPivotCache As Object
    
Call RES

    Sheets("UPH Order").Select
    
    Range("A1:Z1").AutoFilter
    Range("A1:Z1").AutoFilter
    
    Range("A4:Z1000000").Select
    Selection.ClearContents
   

     Set Db = New Connection
     Db.CursorLocation = adUseClient
         
If Db.State = 1 Then Db.Close
        
Db.Open "Provider =IBMDASQL.DataSource.1" & _
        ";Catalog Library List=JDETSTDTA" & _
        ";Persist Security Info=True" & _
        ";Force Translate=0" & _
        ";Data Source = WFVNPROD" & _
        ";User ID =JEFFTRAN" & _
        ";Password =abc111"
        
Call MakeString
            
x = Worksheets("Setting").Range("C6").Text
y = Worksheets("Setting").Range("C7").Text
Z = Worksheets("Setting").Range("F4").Text
     
Application.Calculation = xlCalculationManual
     Worksheets("UPH Order").Select
     Set adors = New Recordset
     If adors.State = 1 Then adors.Close
     
     cmdtxt = "SELECT MOMAST.STID, MOMAST.ORDNO, MOMAST.FITEM, MOMAST.ORQTY+MOMAST.QTDEV as QTY, MOMAST.QTYRC, " & _
              "MOMAST.ODUDT, " & _
              "MOMAST.OSTAT, MOMAST.ITCL, MOMAST.JOBNO, " & _
              "MOMAST.CRDT, MOMAST.CRUS " & _
              "FROM G20ACF9V.AMFLIBW.MOMAST MOMAST " & _
              "WHERE SUBSTR(MOMAST.ORDNO,1,2) In ('MX') " & _
              "AND MOMAST.ORQTY+MOMAST.QTDEV <>0 " & _
              "AND MOMAST.OSTAT in ('10','40','45') AND  MOMAST.FITEM NOT LIKE '%VN%' " & _
              "AND MOMAST.ODUDT between '" & x & "' and '" & y & "' " & _
              "ORDER BY MOMAST.ORDNO "
     
     adors.Open cmdtxt, Db, 3, 3
     

    Worksheets("UPH Order").Range("A2").CopyFromRecordset adors

'Call FFor

Sheets("UPH Order").Select
endrow = Range("B1000000").End(xlUp).Row
Range(Cells(2, 12), Cells(endrow, 25)).FillDown
Application.Calculation = xlCalculationAutomatic

'ENDROW = Range("B1000000").End(xlUp).Row
'Range("L5:W" & ENDROW).Copy
'Range("L5:W" & ENDROW).PasteSpecial (xlPasteValues)


endrow = Range("B1000000").End(xlUp).Row
Range(Cells(endrow + 1, 1), Cells(1000000, 25)).ClearContents

Range("AB4").Select
    ActiveSheet.PivotTables("PivotTable1").PivotCache.Refresh
Range("AK4").Select
    ActiveSheet.PivotTables("PivotTable2").PivotCache.Refresh
Range("AU3").Select
    ActiveSheet.PivotTables("PivotTable3").PivotCache.Refresh

Sheets("Setting").Select
 
End Sub