Sub maketring()

Dim skus As String, i&, arr, aRes 

Sheet1.Activate
arr = Sheet1.Range("a2:a" & Range("a1048576").End(3).Row)

    Set d1 = CreateObject("scripting.dictionary")
    d1.CompareMode = vbTextCompare
    For k = 1 To UBound(arr)
       If Not d1.exists(arr(k, 1)) Then
        d1(arr(k, 1)) = ""
        End If
    Next
    aRes = d1.keys
    
    For i = 0 To UBound(aRes)
        If i < UBound(aRes) Then
            item = aRes(i)
            skus = "'" & item & "'" & "," & skus
        Else
            item = aRes(i)
            skus = skus & "'" & item & "'"
        End If
    Next

Debug.Print (skus)

End Sub