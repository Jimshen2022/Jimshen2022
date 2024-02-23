Sub RPCPC1vs1020Compared()
    
    t = Timer
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
   
    Call PullAS400OnHand
    Call Pull102026nHand
    Call PullASYard
    Call PullAdjustment
    Call AddColumnAFor10200126
    Call Sheet3ColumnHtoL
    
    Application.Calculation = xlCalculationAutomatic
    ThisWorkbook.Save
    Application.ScreenUpdating = True
    MsgBox "Updated Successful~    " & Format(Timer - t, "0.00" & "s")
    
    
End Sub
