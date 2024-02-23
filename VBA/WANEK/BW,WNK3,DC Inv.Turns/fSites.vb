' Dowload 368 找SN loading Door, 再判断Sites
' D9* --- BW
' D8* --- DC
' D4* --- WN3 

Sub SitsNew() 
     't = Timer
    Application.ScreenUpdating = False
    Dim wb As Workbook
    Dim arr, brr, i&, d As Object
    
    
    Set wb = GetObject("C:\Users\jishen\Downloads\WN3_368_Trx.xlsx")    
    arr = wb.ActiveSheet.[a1].CurrentRegion
    wb.Close False
       
    Set d = CreateObject("scripting.dictionary")
    d.CompareMode = vbTextCompare 
    
    For i = 2 To UBound(arr)
        d(arr(i, 5)) = arr(i, 9)
    Next
    
    Sheet5.Activate
    
    brr = Sheet5.Range("a1").CurrentRegion
    
    For i = 2 To UBound(brr)
        If d(brr(i, 13)) Like "D4*" Then
            brr(i, 1) = "Wanek3"
        ElseIf d(brr(i, 13)) Like "D8*" Then
            brr(i, 1) = "DC"
        ElseIf d(brr(i, 13)) Like "D9*" Then
            brr(i, 1) = "BW"
        Else: brr(i, 1) = brr(i, 1)
        End If
        
    Next
    
    With Sheet5.Range("a1").CurrentRegion

        .Columns("e:aj").NumberFormat = "@"
        .Range("a1").Resize(UBound(brr), UBound(brr, 2)).Value = brr
        .Value = brr
    End With
    Erase arr
    Erase brr
   
    Application.ScreenUpdating = True
    'MsgBox "Updated Successful~    " & Format(Timer - t, "0.00" & "s")
End Sub
