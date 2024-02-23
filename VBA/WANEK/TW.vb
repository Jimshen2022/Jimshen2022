Sub generatemac()
t = Timer
Application.ScreenUpdating = False
'Application.Calculation = xlCalculationManual
Application.StatusBar = "Calculating, please wait ......"

Worksheets("Mac file").Range("a1:iv65536").Clear

Dim i As Integer, j As Integer, n As Integer, m As Integer, newtxt As Object, fso As Object

n = Worksheets("MO").Range("a65536").End(xlUp).Row
Worksheets("Mapics Macro").Range("a1:a3").Copy Destination:=Worksheets("Mac file").Range("a1")
m = Worksheets("Mac file").Range("a65536").End(xlUp).Row

For i = 2 To n
If Cells(i, 1) = "" Then Exit Sub

     Worksheets("Mapics Macro").Range("a5") = """" & Worksheets("MO").Cells(i, 1)
     Worksheets("Mapics Macro").Range("a7") = """" & Worksheets("MO").Cells(i, 2)
     Worksheets("Mapics Macro").Range("a9") = """" & Worksheets("MO").Cells(i, 3)
     Worksheets("Mapics Macro").Range("a12") = """" & Worksheets("MO").Cells(i, 4)
     Worksheets("Mapics Macro").Range("a14") = """" & Worksheets("MO").Cells(i, 5)
     Worksheets("Mapics Macro").Range("a16") = """" & Worksheets("MO").Cells(i, 6)
     Worksheets("Mapics Macro").Range("a20") = """" & Worksheets("MO").Cells(i, 7)
     Worksheets("Mapics Macro").Range("a22") = """" & Worksheets("MO").Cells(i, 8)
     Worksheets("Mapics Macro").Range("a5:a28").Copy Destination:=Worksheets("Mac file").Range("a" & m + 2)
     m = Worksheets("Mac file").Range("a65536").End(xlUp).Row

Next i


Set fso = CreateObject("Scripting.FileSystemObject")
Set newtxt = fso.createtextfile(ThisWorkbook.Path & "\TW" & Format(Now(), "mmdd.hhmm") & ".mac")

For j = 1 To m
newtxt.writeline (Worksheets("Mac file").Range("a" & j))
Next j

'Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.StatusBar = False
MsgBox "Finished~~ " & Format(Timer - t, "0.00" & "s")

End Sub
