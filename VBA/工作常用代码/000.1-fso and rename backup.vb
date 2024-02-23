
Sub SaveSheetsForBI()
    'On Error Resume Next
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim wksSht As Worksheet

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")     ' fsoÓÃÓÚadd,edit,move,copy,delete ÎÄ¼þÓëÎÄ¼þ¼Ð
    
    For Each wksSht In Worksheets
        If wksSht.Name = "InvTurnOverRatio" Then
           'wksSht.Range("a9:e12").Cut Destination:=wksSht.Range("g1")
           wksSht.Copy
           wksSht.Columns("a:k").AutoFit
           'ActiveWorkbook.SaveAs filename:="X:\BW_MPT_INV_TURNOVER_RATIO.xlsx"
            wksSht.Copy
           ActiveWorkbook.SaveAs filename:="D:\Document\03-Wanek3\00-Report\1.Inv. turns\BI\BW_MPT_INV_TURNOVER_RATIO.xlsx"
           ActiveWorkbook.Close
        End If
    Next wksSht
    
    fso.CopyFile "D:\Document\03-Wanek3\00-Report\1.Inv. turns\BI\BW_MPT_INV_TURNOVER_RATIO.xlsx", "D:\Document\03-Wanek3\00-Report\1.Inv. turns\BI\BwMptTurnOverBackup\"
    Name "D:\Document\03-Wanek3\00-Report\1.Inv. turns\BI\BwMptTurnOverBackup\BW_MPT_INV_TURNOVER_RATIO.xlsx" _
        As "D:\Document\03-Wanek3\00-Report\1.Inv. turns\BI\BwMptTurnOverBackup\" & "BW_MPT_INV_TURNOVER_RATIO-" & Format(Now(), "yyyy.mm.dd.hhmm") & ".xlsx"
    'ÉÏ2ÐÐµÚÒ»¸öNameÓÃÓÚ¸ÄÎÄ¼þÃû£¬»òÎÄ¼þ¼Ð
    Set wksSht = Nothing
Application.ScreenUpdating = True
Application.DisplayAlerts = True

MsgBox "Saved sheet 'InvTurnOverRatio' to D:\Document\03-Wanek3\00-Report\1.Inv. turns\BI\BW_MPT_INV_TURNOVER_RATIO.xlsx  "
End Sub