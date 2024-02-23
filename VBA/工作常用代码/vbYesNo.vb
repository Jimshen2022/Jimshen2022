
Sub TEST11()
    
    a = MsgBox("Are you sure?", vbYesNo)
    If a = vbNo Then Exit Sub
    
    If a = vbYes Then GoTo 100
    
100
    
    MsgBox "test"
End Sub
