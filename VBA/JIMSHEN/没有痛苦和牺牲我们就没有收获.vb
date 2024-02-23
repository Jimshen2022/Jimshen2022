'Without pain, without sacrifice, we would have Nothing.
'没有痛苦和牺牲 ，我们就没有收获 。

Sub wanek_ctn_utilization()
    
    t = Timer
    ActiveWorkbook.SaveAs ActiveWorkbook.Path & "\wanek_container_utilization -" & Format(Now(), "yyyymmdd.hhmm") & ".xlsm"
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Calculating, please wait ......"
    Dim app As New Application
    
    Call unfilter3
    Call Pull_Wanek_362_trx
    Call data_columnAtoI
    'Call refresh_piovtTables
    
    Sheet6.Select
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Sheet1.Select
    ThisWorkbook.Save
    
    app.Speech.Speak "Dear master, the report form has been completed, please review it"
    MsgBox "Updated Successful~    " & Format(Timer - t, "0.00" & "s")
End Sub
