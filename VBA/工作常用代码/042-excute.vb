Sub wanek3_turn_over_rate_report()
    
    t = Timer
    'ActiveWorkbook.SaveAs ActiveWorkbook.Path & "\Ashton RP Open Orders Fulfillment-" & Format(Now(), "yyyymmdd.hhmm") & ".xlsm"
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Calculating, please wait ......"
    
    
    Call PullWanek3STO
    Call Pull_Wanek3_DC_361_OUT_trx
    Call Pull_Wanek3_111_in_trx

    Sheet6.Select
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    ThisWorkbook.Save
	MsgBox "Updated Successful " & Chr(10) & " Udpated took " & Format(Timer - t, "#,##.00") & "s."
    
    
End Sub