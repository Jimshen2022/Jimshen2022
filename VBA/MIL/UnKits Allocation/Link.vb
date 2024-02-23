Sub ALLOCATE()
Application.ScreenUpdating = False
Dim i As Integer, j As Integer, m As Integer, n As Integer, L As Integer
Dim CO As String
Dim TCO As String
Dim WCO As String
Dim QCO As Integer


Worksheets("LINK").Range("A2:M665536").ClearContents
Worksheets("MO").Range("H2:L665536").ClearContents

m = Worksheets("PO_List").Range("A665536").End(xlUp).Row
n = Worksheets("MO").Range("A665536").End(xlUp).Row

For y = 2 To n
    Worksheets("MO").Range("H" & y) = 0
    Worksheets("MO").Range("I" & y) = Worksheets("MO").Range("E" & y) - Worksheets("MO").Range("H" & y)
Next y
    
For i = 5 To m
    CO = Worksheets("PO_List").Range("A" & i)
    TCO = Worksheets("PO_List").Range("B" & i)
    WCO = Worksheets("PO_List").Range("C" & i)
    QCO = Worksheets("PO_List").Range("D" & i)
    
    L = Worksheets("LINK").Range("a65536").End(xlUp).Row
    Debug.Print L
    
    Worksheets("LINK").Range("A" & L + 1) = CO
    Worksheets("LINK").Range("B" & L + 1) = TCO
    Worksheets("LINK").Range("C" & L + 1) = WCO
    Worksheets("LINK").Range("D" & L + 1) = QCO
    
 TMPMOQTY = 0
    
    Set x = CreateObject("ADODB.Connection")
    x.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties='Excel 8.0;hdr=YES;';Data Source=" & ThisWorkbook.FullName
    SQL = "SELECT MO, FG, MO_BAL,"""" , DUEDATE FROM [MO$] WHERE FG = """ & TCO & """ AND MO_BAL>0 ORDER by DUEDATE "
    
    Set y = x.Execute(SQL)
    Worksheets("LINK").Range("E" & L + 1).CopyFromRecordset y
    x.Close
      
   LE = Worksheets("LINK").Range("E665536").End(xlUp).Row

   For j = L + 1 To LE
      TMPCOQTY = Worksheets("LINK").Range("D" & j)
      TMPMOQTY = TMPMOQTY + Worksheets("LINK").Range("G" & j)
      
      If TMPCOQTY <= TMPMOQTY Then
         Worksheets("LINK").Range("H" & j) = TMPCOQTY
         Worksheets("LINK").Range("G" & j) = TMPMOQTY - TMPCOQTY
         Exit For
      Else
         Worksheets("LINK").Range("H" & j) = Worksheets("LINK").Range("G" & j)
         Worksheets("LINK").Range("G" & j) = 0
         Worksheets("LINK").Range("D" & j + 1) = Worksheets("LINK").Range("D" & j) - Worksheets("LINK").Range("H" & j)
         Worksheets("LINK").Range("A" & j + 1) = Worksheets("LINK").Range("A" & j)
         Worksheets("LINK").Range("B" & j + 1) = Worksheets("LINK").Range("B" & j)
         Worksheets("LINK").Range("C" & j + 1) = Worksheets("LINK").Range("C" & j)
         Worksheets("LINK").Range("D" & j) = Worksheets("LINK").Range("H" & j)

         TMPMOQTY = Worksheets("LINK").Range("G" & j)
      End If
  Next j
  
 Set dic = CreateObject("SCRIPTING.DICTIONARY")
     
           Rng = Worksheets("LINK").Range("E1:H" & LE)
           For r = 2 To UBound(Rng)
               y = Rng(r, 1)
               dic(y) = dic(y) + Rng(r, 4)
           Next r
     
   Worksheets("MO").Select
   Tarr = Range(Cells(1, 1), Cells(n, 9))
   
   For r = 2 To UBound(Tarr)
            y = Trim(Tarr(r, 4))
            Tarr(r, 8) = dic(y)
            If Tarr(r, 8) = "" Then
               Tarr(r, 8) = 0
            End If
            Tarr(r, 9) = Tarr(r, 5) - Tarr(r, 8)
   Next r
   Range(Cells(1, 1), Cells(n, 9)) = Tarr
   Worksheets("LINK").Select
    
Next i
L = Worksheets("LINK").Range("a665536").End(xlUp).Row
LE = Worksheets("LINK").Range("E665536").End(xlUp).Row

If LE > L Then
   Range("E" & L + 1 & ":H" & LE).ClearContents
End If

Application.ScreenUpdating = True

    Range("J2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(ROUND(1/VLOOKUP(RC2,MatInChar!C3:C8,6,0),0),0)"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "=ROUNDUP(RC[-3]/RC[-1],0)"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(VLOOKUP(RC[-7],UPHMO!C3:C13,11,0)=""A"",""Day Shift"",IF(VLOOKUP(RC[-7],UPHMO!C3:C13,11,0)=""B"",""Night Shift"",""Day Shift"")),""Day Shift"")"
    
    
    endrow = Worksheets("LINK").Range("A65000").End(xlUp).Row
    If endrow < 2 Then
    endrow = 2
    End If
    
    Range("J2:L2").Copy
    Range("J3:L" & endrow).PasteSpecial (xlPasteFormulas)

    Calculate
    
    'Range("J2:L" & endrow).Select
    'Selection.Copy
    'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'Application.CutCopyMode = False
    
End Sub

