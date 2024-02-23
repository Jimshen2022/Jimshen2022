Public Sub Load_FinalMO()

Dim i As Long
Dim adors As New Recordset
Application.Calculation = xlManual
Worksheets("UPHMO").Range("A3:q50000").ClearContents
    
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
     

     Set adors = New Recordset
     If adors.State = 1 Then adors.Close
          
     cmdtxt = "SELECT A.FITWH, A.REFNO, A.ORDNO, A.FITEM, A.FDESC, " & _
            "A.ORQTY + A.QTDEV -A.QTYRC as MQTY,A.QTYRC, " & _
            "Date(Substr(Char(A.ODUDT+ 19000000), 1, 4) || '-'||  Substr(Char(A.ODUDT + 19000000), 5, 2)|| '-' ||substr(Char(A.ODUDT + 19000000), 7, 2)) AS FG_DUE, " & _
            "A.OSTAT, A.JOBNO, A.ITCL " & _
            "FROM AMFLIBL.MOMAST A " & _
            "WHERE (A.FITWH='51') AND (substr(A.ORDNO,1,2)='MA') " & _
            "AND (SUBSTR(A.JOBNO, 12, 1) NOT IN ('O','S','P')) " & _
            "AND (A.OSTAT Not In ('99','45','55')) AND (A.ORQTY + A.QTDEV -A.QTYRC <>0) "
            
     adors.Open cmdtxt, Db, 3, 3
     
     For i = 0 To adors.Fields.Count - 1
         Worksheets("UPHMO").Cells(2, i + 1) = adors.Fields(i).Name
     Next i
     
     Worksheets("UPHMO").Range("A3").CopyFromRecordset adors
     
     adors.Close
     Set adors = Nothing

Set sht = Sheets("UPHMO")

LastRow = sht.Cells(sht.Rows.Count, "A").End(xlUp).Row

'Shift
If LastRow = 2 Then
LastRow = 3
End If

For y = 3 To LastRow

    sht.Range("M" & y) = Mid(sht.Range("J" & y), 2, 1)
    sht.Range("L" & y) = Mid(sht.Range("J" & y), 1, 4)
    'sht.Range("O" & y) = Application.WorksheetFunction.VLookup(Mid(sht.Range("L" & y), 1, 1), Worksheets("Setting").Range("M2:N100"), 2, 0)
    sht.Range("P" & y) = Mid(sht.Range("J" & y), 12, 1)

If Mid(sht.Range("D" & y), 1, 1) = "U" Then
    sht.Range("N" & y) = Mid(sht.Range("D" & y), 1, 4)
    
Else
    sht.Range("N" & y) = Mid(sht.Range("D" & y), 1, 3)
End If

Next y

    ActiveWorkbook.Worksheets("UPHMO").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("UPHMO").Sort.SortFields.Add Key:=Range( _
        "H3:H" & LastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("UPHMO").Sort.SortFields.Add Key:=Range( _
        "M3:M" & LastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("UPHMO").Sort.SortFields.Add Key:=Range( _
        "L3:L" & LastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("UPHMO").Sort
        .SetRange Range("A2:P" & LastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Call Sumarry1st
MsgBox "UPHOrder Finished"
End Sub

