Sub Sumarry1st()

Dim i As Integer
Dim j As Integer
Dim b As Integer
Set ws = Sheets("Onhand")
Set pr = Sheets("MO")
Application.Calculation = xlManual
    
    If Worksheets("MO").Range("A1") <> "" Then
    Worksheets("MO").Range("A1").AutoFilter Field:=1
    Worksheets("MO").Range("A1").AutoFilter
    End If
    
LR = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

StartTime = Timer
pr.Range("A2:I10000").ClearContents      'clear old data

j = 2

For i = 2 To LR

pr.Cells(j, 1) = ws.Range("A" & i)
pr.Cells(j, 2) = Worksheets("Setting").Range("A" & 4)
pr.Cells(j, 3) = ws.Range("B" & i)
pr.Cells(j, 4) = "MX" & (i - 1)
pr.Cells(j, 5) = ws.Range("E" & i)

j = j + 1

Next i

endr = pr.Cells(pr.Rows.Count, "A").End(xlUp).Row
Debug.Print endr
Call Sumarry2nd

pr.Select
Columns("B:B").Select
    Selection.NumberFormat = "General"
    
End Sub

Sub Sumarry2nd()
Dim i As Long
Dim x As Integer
Dim y As Integer
Dim z As Integer

Set sht = Sheets("UPHMO")
Set fisht = Sheets("MO")

LastRow = sht.Cells(sht.Rows.Count, "A").End(xlUp).Row
endr = fisht.Cells(fisht.Rows.Count, "A").End(xlUp).Row

'Shift
If LastRow = 3 Then
LastRow = 3
End If
x = endr + 1
For y = 3 To LastRow
   
   
    fisht.Range("A" & x) = sht.Range("A" & y)
    fisht.Range("B" & x) = sht.Range("H" & y)
    fisht.Range("C" & x) = sht.Range("D" & y)
    fisht.Range("D" & x) = sht.Range("C" & y)
    fisht.Range("E" & x) = sht.Range("F" & y)
    fisht.Range("F" & x) = sht.Range("E" & y)
    fisht.Range("G" & x) = sht.Range("J" & y)
   
x = x + 1


Next y

Sheets("MO").Select
Columns("B:B").Select
    Selection.TextToColumns Destination:=Range("B1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(0, 3), TrailingMinusNumbers:=True
        
End Sub

