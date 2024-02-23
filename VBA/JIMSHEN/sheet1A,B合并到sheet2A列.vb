' D:\Document\00-SQL\Notepad\VBA\excelfiles\Jim tested.xlsb  

Sub fsj()
Dim i, j, k As Long
Dim m As Integer

Sheet2.Range("A2:H1000000").ClearContents

For m = 2 To 545

k = Sheet2.Range("A1048576").End(xlUp).row + 1

    i = Sheet1.Range("B" & m) - Sheet1.Range("A" & m) + 1
    For j = 1 To i
        Sheet2.Range("A" & k + j - 1) = Format(Sheet1.Range("A" & m) + j - 1, "00000")
    Next
    Sheet1.Range("C" & m & ":I" & m).Copy Sheet2.Range("B" & k & ":H" & k + j)

Next

End Sub

