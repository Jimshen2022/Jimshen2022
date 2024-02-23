Sub BW_Inv_Turns()

Application.ScreenUpdating = False
Dim arr, ii&, brr, d As Object
arr = Sheet3.Range("a1").CurrentRegion

Set d = CreateObject("scripting.dictionary")
    d.CompareMode = vbTextCompare '²»Çø·Ö×ÖÄ¸´óÐ¡Ð´

For ii = 1 To UBound(arr)
    d(arr(ii, 1)) = d(arr(ii, 1)) + arr(ii, 3)
Next


'Date
Dim i&, x&, y&, nrow&, z&, x1&, y1&, z1&, r&
    Worksheets("BW Turns").Select
    With Worksheets("BW Turns")
        x = .Range("b1048576").End(3).Value
        y = Date
        nrow = .Range("a1048576").End(3).row
        For i = x + 1 To y
            For j = 1 To y - x
                .Range("b" & nrow + j) = i
               i = i + 1
            Next j
        Next i
        
        'Week
        z = .Range("b1048576").End(3).row
        For i = nrow To z - 1
            .Range("a" & i + 1).Value = Application.WeekNum(.Range("b" & i + 1))
        Next
        
        'C to f Columns
        
        r = .Range("c1048576").End(3).row
        nrow = .Range("a1048576").End(3).row
        .Range("c" & r - 1, "F" & r - 1).AutoFill Destination:=.Range("c" & r - 1, "F" & nrow)
        
            
    Worksheets("BW Turns").Range("E1048576").End(3).Value = d("BW")
    End With
    
    Calculate
    Erase arr
    Set d = Nothing
    
Application.ScreenUpdating = True
End Sub


Sub WN3_Inv_Turns()

Application.ScreenUpdating = False
Dim arr, ii&, brr, d As Object
arr = Sheet3.Range("a1").CurrentRegion

Set d = CreateObject("scripting.dictionary")
    d.CompareMode = vbTextCompare '²»Çø·Ö×ÖÄ¸´óÐ¡Ð´

For ii = 1 To UBound(arr)
    d(arr(ii, 1)) = d(arr(ii, 1)) + arr(ii, 3)
Next


'Date
Dim i&, x&, y&, nrow&, z&, x1&, y1&, z1&, r&
    Worksheets("WNK3 Turns").Select
    With Worksheets("WNK3 Turns")
        x = .Range("b1048576").End(3).Value
        y = Date
        nrow = .Range("a1048576").End(3).row
        For i = x + 1 To y
            For j = 1 To y - x
                .Range("b" & nrow + j) = i
               i = i + 1
            Next j
        Next i
        
        'Week
        z = .Range("b1048576").End(3).row
        For i = nrow To z - 1
            .Range("a" & i + 1).Value = Application.WeekNum(.Range("b" & i + 1))
        Next
        
        'C to f Columns
        
        r = .Range("c1048576").End(3).row
        nrow = .Range("a1048576").End(3).row
        .Range("c" & r - 1, "F" & r - 1).AutoFill Destination:=.Range("c" & r - 1, "F" & nrow)
        
            
    Worksheets("WNK3 Turns").Range("E1048576").End(3).Value = d("Wanek3")
    End With
    
    Calculate
    Erase arr
    Set d = Nothing
    
Application.ScreenUpdating = True
End Sub


Sub DC_Inv_Turns()

Application.ScreenUpdating = False
Dim arr, ii&, brr, d As Object
arr = Sheet3.Range("a1").CurrentRegion

Set d = CreateObject("scripting.dictionary")
    d.CompareMode = vbTextCompare '²»Çø·Ö×ÖÄ¸´óÐ¡Ð´

For ii = 1 To UBound(arr)
    d(arr(ii, 1)) = d(arr(ii, 1)) + arr(ii, 3)
Next


'Date
Dim i&, x&, y&, nrow&, z&, x1&, y1&, z1&, r&
    Worksheets("DC Turns").Select
    With Worksheets("DC Turns")
        x = .Range("b1048576").End(3).Value
        y = Date
        nrow = .Range("a1048576").End(3).row
        For i = x + 1 To y
            For j = 1 To y - x
                .Range("b" & nrow + j) = i
               i = i + 1
            Next j
        Next i
        
        'Week
        z = .Range("b1048576").End(3).row
        For i = nrow To z - 1
            .Range("a" & i + 1).Value = Application.WeekNum(.Range("b" & i + 1))
        Next
        
        'C to f Columns
        
        r = .Range("c1048576").End(3).row
        nrow = .Range("a1048576").End(3).row
        .Range("c" & r - 1, "F" & r - 1).AutoFill Destination:=.Range("c" & r - 1, "F" & nrow)
        
            
    Worksheets("DC Turns").Range("E1048576").End(3).Value = d("DC")
    End With
    
    Calculate
    Erase arr
    Set d = Nothing
    
Application.ScreenUpdating = True
End Sub
