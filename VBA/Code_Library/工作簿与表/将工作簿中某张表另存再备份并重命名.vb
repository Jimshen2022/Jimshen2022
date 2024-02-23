
' Wanek3 inventory turnover ratio 
'D:\Document\03-Wanek3\00-Report\1.Inv. turns\BW_DC_WN3 Inventory Turns Version01.xlsb

' 将工作簿中某张表另存，再备份后重命名




'MIL INVENTORY TURNS BEFORE CLOSE THE FILE FOR BACKUP AND COPY FOR BI report
'D:\Document\06-Millennium\00-Report\01-InvTurns

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    Application.DisplayAlerts = False
    Application.DisplayAlerts = False
    Dim wksSht As Worksheet
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ans = MsgBox("Do you want to copy Product for BI report? ", vbYesNo)
    
    If ans = vbYes Then
        For Each wkSht In Worksheets
            If wksSht.Name = "Product" Then
                wksSht.Copy
                ActiveWorkbook.SaveAs Filename:="D:\Document\06-Millennium\00-BI\MIL Product Seperated-2099.xlsx"
                ActiveWorkbook.Close
        'fso.CopyFile "D:\Document\03-Wanek3\00-Report\1.Inv. turns\BI\BW_MPT_INV_TURNOVER_RATIO.xlsx", "D:\Document\03-Wanek3\00-Report\1.Inv. turns\BI\BwMptTurnOverBackup\"
            End If
        Next wkSht
    Else:
        Exit Sub
    End If
        
    ans2 = MsgBox("Do you want to backup the file ?", vbYesNo)
    If ans2 = vbYes Then
        ThisWorkbook.SaveAs ("D:\Document\02-Ashton\00-Report\04-Inv. Turns\ATInvTurnsBackup\Ashton Inventory Turns - " & Format(Now(), "yyyymmdd.hhmm") & ".xlsm")
        Else:
        Exit Sub
    End If

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub













Sub SaveSheets()
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim wksSht As Worksheet
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")     ' fsoÓÃÓÚadd,edit,move,copy,delete ÎÄ¼þÓëÎÄ¼þ¼Ð
    
    For Each wksSht In Worksheets
        If wksSht.Name = "InvTurnOverRatio" Then
           wksSht.Range("a9:e12").Cut Destination:=wksSht.Range("g1")
           wksSht.Columns("a:k").AutoFit
           wksSht.Copy
           ActiveWorkbook.SaveAs filename:="D:\Document\03-Wanek3\00-Report\1.Inv. turns\BI\BW_MPT_INV_TURNOVER_RATIO.xlsx"
           ActiveWorkbook.Close
        End If
    Next wksSht
    
    fso.CopyFile "D:\Document\03-Wanek3\00-Report\1.Inv. turns\BI\BW_MPT_INV_TURNOVER_RATIO.xlsx", "D:\Document\03-Wanek3\00-Report\1.Inv. turns\BI\BwMptTurnOverBackup\"
    Name "D:\Document\03-Wanek3\00-Report\1.Inv. turns\BI\BwMptTurnOverBackup\BW_MPT_INV_TURNOVER_RATIO.xlsx" _
        As "D:\Document\03-Wanek3\00-Report\1.Inv. turns\BI\BwMptTurnOverBackup\" & "BW_MPT_INV_TURNOVER_RATIO-" & Format(Now(), "yyyy.mm.dd.hhmm") & ".xlsx"

    Set wksSht = Nothing
Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub





Private Sub Workbook_BeforeClose(Cancel As Boolean)
Application.DisplayAlerts = False
ans = MsgBox("Do you want to move the file to Ashton file server? ", vbYesNo)
On Error Resume Next
If ans = vbYes Then

    On Error Resume Next
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim wksSht As Worksheet
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")     ' fsoÓÃÓÚadd,edit,move,copy,delete ÎÄ¼þÓëÎÄ¼þ¼Ð
    
    For Each wksSht In Worksheets
        If wksSht.Name = "Inv.Turns" Then
           wksSht.Copy
           ActiveWorkbook.SaveAs Filename:="\\10.141.100.133\AshtonData\UPH FG Warehouse\Public\Inventory\Inventory-Jim\InventoryTurnsForSam\Ashton_InvTurnOver_Ratio.xlsx"
           ActiveWorkbook.Close
        End If
    Next wksSht
    Set wksSht = Nothing
End If

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub





Private Sub Workbook_BeforeClose(Cancel As Boolean)
Application.DisplayAlerts = False
ans = MsgBox("Do you want to move the file to Ashton file server? ", vbYesNo)
    On Error Resume Next
    If ans = vbYes Then
    
        On Error Resume Next
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")     ' fsoÓÃÓÚadd,edit,move,copy,delete ÎÄ¼þÓëÎÄ¼þ¼Ð
        fso.CopyFile "D:\Document\02-Ashton\00-Report\00-Weekly report\Ashton Inv. weekly report - 2022.xlsx", "\\10.141.100.133\AshtonData\UPH FG Warehouse\Public\Inventory\Inventory-Jim\InventoryTurnsForSam\"
    End If

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

