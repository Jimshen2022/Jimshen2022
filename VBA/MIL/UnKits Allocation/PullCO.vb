Sub MakeString()
Dim i As Integer
Dim j As Integer

Sheets("SETTING").Select

'Etd
If Sheets("Setting").Range("b" & 6).Value = "" Then Range("a" & 1).Value = ""

If Sheets("Setting").Range("b" & 6).Value <> "" Then
Range("a" & 1).Value = "(" & "'" & Sheets("SETTING").Range("b6").Value & "'"
For i = 7 To 700
If Sheets("SETTING").Range("b" & i).Value = "" Then Exit For
Range("a" & 1).Value = Range("a" & 1).Value & "," & "'" & Sheets("SETTING").Range("b" & i).Value & "'"
Next i
Range("a" & 1).Value = Range("a" & 1).Value & ")"
End If

'Sku
If Sheets("Setting").Range("c" & 6).Value = "" Then Range("b" & 1).Value = ""

If Sheets("Setting").Range("c" & 6).Value <> "" Then
Range("b" & 1).Value = "(" & "'" & Sheets("SETTING").Range("c6").Value & "'"
For i = 7 To 700
If Sheets("SETTING").Range("c" & i).Value = "" Then Exit For
Range("b" & 1).Value = Range("b" & 1).Value & "," & "'" & Sheets("SETTING").Range("c" & i).Value & "'"
Next i
Range("b" & 1).Value = Range("b" & 1).Value & ")"
End If

'Itcls
If Sheets("Setting").Range("d" & 6).Value = "" Then Range("c" & 1).Value = ""

If Sheets("Setting").Range("d" & 6).Value <> "" Then
Range("c" & 1).Value = "(" & "'" & Sheets("SETTING").Range("d6").Value & "'"
For i = 7 To 700
If Sheets("SETTING").Range("d" & i).Value = "" Then Exit For
Range("c" & 1).Value = Range("c" & 1).Value & "," & "'" & Sheets("SETTING").Range("d" & i).Value & "'"
Next i
Range("c" & 1).Value = Range("c" & 1).Value & ")"
End If


'whs
If Sheets("Setting").Range("e" & 6).Value = "" Then Range("d" & 1).Value = ""

If Sheets("Setting").Range("e" & 6).Value <> "" Then
Range("d" & 1).Value = "(" & "'" & Sheets("SETTING").Range("e6").Value & "'"
For i = 7 To 700
If Sheets("SETTING").Range("e" & i).Value = "" Then Exit For
Range("d" & 1).Value = Range("d" & 1).Value & "," & "'" & Sheets("SETTING").Range("e" & i).Value & "'"
Next i
Range("d" & 1).Value = Range("d" & 1).Value & ")"
End If
      
'Cont
If Sheets("Setting").Range("f" & 6).Value = "" Then Range("f" & 1).Value = ""

If Sheets("Setting").Range("f" & 6).Value <> "" Then
Range("f" & 1).Value = "(" & "'" & Sheets("SETTING").Range("f6").Value & "'"
For i = 7 To 700
If Sheets("SETTING").Range("f" & i).Value = "" Then Exit For
Range("f" & 1).Value = Range("f" & 1).Value & "," & "'" & Sheets("SETTING").Range("f" & i).Value & "'"
Next i
Range("f" & 1).Value = Range("f" & 1).Value & ")"
End If

End Sub
Sub CO()
'On Error Resume Next
Application.ScreenUpdating = False
Dim i As Integer, j As Integer, n As Integer, m As Integer
Dim cmdtxt As String
Dim adors As New Recordset
Dim sh As Worksheet
Dim adoCN As Object
Dim strSQL As String
Dim objPivotCache As Object

Sheets("PO_List").Select
If Worksheets("PO_List").Range("A4") <> "" Then
   Worksheets("PO_List").Range("A4").AutoFilter Field:=1
   Worksheets("PO_List").Range("A4").AutoFilter
End If

Sheets("PO_List").Select
Range("A5:G10000").Select
Selection.ClearContents
    
Worksheets("SETTING").Select

Call MakeString
etd = Worksheets("SETTING").Range("a1").Text
sku = Worksheets("SETTING").Range("b1").Text
itcls = Worksheets("SETTING").Range("c1").Text
whs = Worksheets("SETTING").Range("d1").Text
Debug.Print etd
Debug.Print sku
Debug.Print itcls
Debug.Print whs

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
         
     Worksheets("PO_List").Select
     Set adors = New Recordset
     If adors.State = 1 Then adors.Close
     
      cmdtxt = "SELECT A.CDCVNB as OPOrder, A.CDAITX as #SKU, " & _
                "A.CDB9CD as Warehouse, A.CDAGNV as Qty, B.ITCLS,Date(Substr(Char(A.CDD0NB+ 19000000), 1, 4) || '-'||  Substr(Char(A.CDD0NB + 19000000), 5, 2)|| '-' ||substr(Char(A.CDD0NB + 19000000), 7, 2)) AS ETD " & _
                "FROM AMFLIBL.MBCDRESM A, AMFLIBL.ITEMASA B " & _
                "WHERE B.ITNBR = A.CDAITX AND A.CDAGNV > 0 AND A.CDAITX like '%UN%' " & _
                "AND B.ITCLS not in ('ZKIS','ZKIZ') "

        If etd <> "" Then
              cmdtxt = cmdtxt & " AND A.CDD0NB in " & etd & ""
        End If
        
        If whs <> "" Then
              cmdtxt = cmdtxt & " AND A.CDA3CD in " & whs & ""
        End If
                
        If sku <> "" Then
              cmdtxt = cmdtxt & " AND A.CDAITX in " & sku & ""
        End If
        
        If itcls <> "" Then
              cmdtxt = cmdtxt & " AND B.ITCLS in " & itcls & ""
        End If
        
        
        cmdtxt = cmdtxt
        
    
      adors.Open cmdtxt, Db, 3, 3
     
     For i = 0 To adors.Fields.Count - 1
         Worksheets("PO_List").Cells(4, i + 1) = adors.Fields(i).Name
     Next i
      

    Worksheets("PO_List").Range("A5").CopyFromRecordset adors
    
endrow = Worksheets("PO_List").Range("A65000").End(xlUp).Row
    
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "Priority"
    
    Range("E3").Copy
    Range(Cells(5, 5), Cells(endrow, 5)).PasteSpecial (xlPasteFormulas)

Calculate
    
    Range("E5:E" & endrow).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
Worksheets("PO_List").Select
  
    ActiveWorkbook.Worksheets("PO_List").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("PO_List").Sort.SortFields.Add Key:=Range("F4:F" & endrow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("PO_List").Sort.SortFields.Add Key:=Range("E4:E" & endrow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("PO_List").Sort
        .SetRange Range("A4:F" & endrow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Worksheets("Setting").Select
Application.ScreenUpdating = True
MsgBox "PO_List Downloaded!"

End Sub
Sub Get_Group()

    Dim cnn As New ADODB.Connection
    Dim myPath As String
    Dim myTable As String
    Dim SQL As String

Application.ScreenUpdating = False

Sheets("Get_Group").Select
Columns("A:D").Select
Selection.ClearContents

UName = Worksheets("Setting").Range("B2").Text
UPass = Worksheets("Setting").Range("B3").Text

Application.Calculation = xlManual

     Set Db = New Connection
         Db.CursorLocation = adUseClient
         
      If Db.State = 1 Then Db.Close
        
Db.Open "Provider =IBMDASQL.DataSource.1" & _
        ";Catalog Library List=JDETSTDTA" & _
        ";Persist Security Info=True" & _
        ";Force Translate=0" & _
        ";Data Source = WFVNPROD" & _
        ";User ID =WANEKPIC " & _
        ";Password =WANEKPIC "
         
     Worksheets("Get_Group").Select
     Set adors = New Recordset
     If adors.State = 1 Then adors.Close
     
     cmdtxt = "SELECT ITEMBLS1.FFAITX " & _
              "FROM AMFLIBW.ITEMBLS1 ITEMBLS1, AMFLIBW.ITMRVAL0 ITMRVAL0 " & _
              "WHERE ITEMBLS1.FFAITX = ITMRVAL0.ITNOAD AND ITMRVAL0.STIDAD = ITEMBLS1.FFA3CD AND ITEMBLS1.FFA3CD In ('33','35') " & _
              "AND ITEMBLS1.FFAQCD='FA00' AND ITEMBLS1.FFBBVA+ITEMBLS1.FFBIVA<>0 " & _
              "GROUP BY ITEMBLS1.FFAITX " & _
              "ORDER BY ITEMBLS1.FFAITX "
    
     adors.Open cmdtxt, Db, 3, 3
     
     For i = 0 To adors.Fields.Count - 1
         Worksheets("Get_Group").Cells(1, i + 1) = adors.Fields(i).Name
     Next i

    Worksheets("Get_Group").Range("A2").CopyFromRecordset adors
    
    endrow = Range("A65000").End(xlUp).Row
    Sheets("AS").Select
    Range("B7").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Worksheets("Get_Group").Select
    Range(Cells(endrow + 1, 1), Cells(endrow + 1, 1)).PasteSpecial (xlPasteFormulas)
      
       
    ActiveSheet.Range("$A$1:$A$65000").RemoveDuplicates Columns:=1, Header:= _
        xlYes
        
Application.Calculation = xlManual
    
    Range("B2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-1],ITEMBLS1!C[-1]:C[1],2,0),""-"")"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-2],ITEMBLS1!C[-2]:C,3,0),""-"")"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC3,C5:C6,2,0),""-"")"
    
    endrow = Range("A65000").End(xlUp).Row
    
    Range("B2:D2").Copy
    Range(Cells(2, 2), Cells(endrow, 4)).PasteSpecial (xlPasteFormulas)

Application.Calculation = xlAutomatic

ActiveWorkbook.Save

    Range(Cells(2, 2), Cells(endrow, 4)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

ActiveWorkbook.Save
Worksheets("Get_Group").Select
Application.ScreenUpdating = True
MsgBox "Item Master Downloaded!"
End Sub
