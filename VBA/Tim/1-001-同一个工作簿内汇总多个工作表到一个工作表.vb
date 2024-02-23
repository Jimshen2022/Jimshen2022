Sub MultipleWksAccumulation()
    Dim i%, j%, n%
    
    Sheets("汇总").Range("a2:f100000").Clear
    For i = 1 To Sheets.Count
        If Sheets(i).Name <> "汇总" Then
            y = Sheets(i).Range("a1048576").End(3).Row
            x = Sheets("汇总").Range("b1048576").End(3).Row
            Sheets(i).Range("a2:e" & y).Copy Sheets("汇总").Range("b" & x + 1)
            
            'Sheets(i).Range("a2:e" & y).Copy
            'Sheets("汇总").Range("b" & x + 1 & ":f" & x + y - 1).PasteSpecial xlPasteValues

            Sheets("汇总").Range("a" & x + 1 & ":a" & x + y - 1).Value = Sheets(i).Name
        End If
    Next
End Sub
