Sub dates格式()
    
    Sheet1.Activate
    
    Range("a1") = "DataCollectedAt:  " & Format(Now, "HH:MM:SSam/pm,mmm.dd.yyyy")
    Range("A1").Font.Color = -16776961
	
	
	
	
Sub KNQMAN66()

    t = Timer
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Calculating, please wait ......"
    
    'ActiveWorkbook.SaveAs ActiveWorkbook.Path & "\AT Mapics vs HJ vs KNQMAN Compared Report -" & Format(Now(), "yyyymmdd.hhmm") & ".xlsx"

    Call PullVarianceReport              'Module1 读取HJvsMapicsvsKNQMAN report, 将 knq variance report的文本数字改为数值
    Call delet_variance                  '将report的HJ,MAPICS,KNQ 为0的行删除
    
    Sheet10.Select
    Sheet10.Range("e2:e21").Value = Date     'Date 表示今天
    
    Sheet6.Select
    Sheet6.Range("A2") = "DataCollectedAt:  " & Format(Now, "HH:MM:SSam/pm,mmm.dd.yyyy")
    Sheet6.Range("A2").Font.color = -16776961
    Sheet8.Range("d:d, g:g,l:l").Font.color = -16776961  '不连续列设为红色
    Sheet8.Range("c:l").HorizontalAlignment = 3          '连续列设为置中对齐 3-置中对齐，1-靠左对齐
    Sheet25.Select
    
    
    Application.Calculation = xlCalculationAutomatic
    ThisWorkbook.Save
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    MsgBox "Udpated~ " & Format(Timer - t, "0.00") & "s"
