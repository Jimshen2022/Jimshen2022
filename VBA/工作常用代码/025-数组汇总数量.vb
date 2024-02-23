Sub RP_Received_qty()
    
    Dim k%, q%, nrow&, Arr()
    
    
    nrow = Sheet5.Range("a1048576").End(3).Row
    Sheet5.Range("ac1") = "Qty"
    Range("ac2:ac20000").ClearContents
    Arr = Sheet5.Range("a1").CurrentRegion
    q = 0
    For k = 2 To nrow
        Arr(k, 29) = Arr(k, 14) * 1
        q = Arr(k, 29) + q
    Next
    Sheet5.Range("a1").Resize(UBound(Arr), UBound(Arr, 2)).Value = Arr
    Sheet6.Range("e2").Value = q
    
End Sub