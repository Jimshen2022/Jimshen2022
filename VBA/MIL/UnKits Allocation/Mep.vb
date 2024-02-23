Dim Nufile As Integer
Dim ifile As Integer
Dim ha As Long

Sub Hade()
    On Error Resume Next
    Kill Range("duongdan").Value & "\*.mac*"
     On Error Resume Next
On Error GoTo 0

End Sub

Sub Ha_runmac()

Dim hfist As Long
Dim hend As Long
Dim newtxt As Object, fso As Object
Dim endrow As Long

Sheets("Macro").Select

endrow = Cells(Rows.Count, 10).End(xlUp).Row

V1 = endrow - 4
'Const ha = 10
If (V1 / Range("O" & 1).Value) - Int(V1 / Range("O" & 1).Value) = 0 Then

Nufile = V1 / Range("O" & 1).Value
Else

Sheets("Macro").Range("p" & 1).Value = Int(V1 / Range("O" & 1).Value) + 1

End If

Call Hade

Sheets("Macro").Select

Range("H5:H50000").Interior.PatternColor = xlNone 'forrr testmacro
Range("H5:H50000").ClearContents

For ifile = 1 To Sheets("Macro").Range("p" & 1).Value

hfist = (ifile - 1) * Range("O" & 1).Value + 5
hend = ifile * Range("O" & 1).Value + 4


Cells(hfist, 8).Interior.ColorIndex = 5 'forrr testmacro
Cells(hfist, 8).Value = "First Macro " & ifile
If hend > endrow Then
hend = endrow
End If

Cells(hend, 8).Interior.ColorIndex = 3 'forrr testmacro
Cells(hend, 8).Value = "End Macro " & ifile
Range(Cells(hfist, 9), Cells(hend, 9)).Value = "X"

Call GetMapicsCodesH

Range(Cells(hfist, 9), Cells(hend, 9)).ClearContents

Next

MsgBox (" Save File at " & Range("duongdan"))

End Sub

Sub GetMapicsCodesH()
Dim i, j, n, m, d As Integer

Application.ScreenUpdating = False
Worksheets("Macro").Range("F5:F20000").Clear
'n = Worksheets("MO").Range("a20000").End(xlUp).Row

'Part I
For i = 5 To Worksheets("Macro").Range("A20000").End(xlUp).Row
'danh dau p1
If Worksheets("Macro").Range("D" & i).Value = "P2" Then
  Worksheets("Macro").Range("F5" & ":" & "F" & i - 1).Value = Worksheets("Macro").Range("E5" & ":" & "E" & i - 1).Value ' copy toan bo part I sang
 Exit For
 End If
Next i

'Part II, part have the change in code by change the value
d = i
n = Worksheets("Macro").Range("A20000").End(xlUp).Row - 1
m = Worksheets("Macro").Range("J20000").End(xlUp).Row
  
For j = 5 To m ' go from Cell I5 to end of range of data
 If Worksheets("Macro").Range("I" & j).Value <> "" Then
  For k = i To n ' go through Part 2
  ' F
  If Worksheets("Macro").Range("B" & k).Value = "F" Then
   Worksheets("Macro").Range("F" & d).Value = Worksheets("Macro").Range("E" & k).Value
   d = d + 1
  End If
  
  ' C
  If Worksheets("Macro").Range("B" & k).Value = "C" Then
   Worksheets("Macro").Range("F" & d).Value = Worksheets("Macro").Range("E" & k).Value & Worksheets("Macro").Cells(j, Worksheets("Macro").Range("C" & k).Value + 9) & """"
   d = d + 1
  End If
  Next k
 End If
 
Next j
'end code

Worksheets("Macro").Range("F" & d).Value = "end sub"

' write to file .mac
'gio = Time
Set fso = CreateObject("Scripting.FileSystemObject")
Set newtxt = fso.createtextfile(Range("duongdan").Value & "\" & Range("tenfile").Value & ifile & ".mac")
For j = 5 To Worksheets("Macro").Range("F20000").End(xlUp).Row
newtxt.writeline (Worksheets("Macro").Range("F" & j))
Next j

'ActiveSheet.Protect

Application.ScreenUpdating = True

End Sub

