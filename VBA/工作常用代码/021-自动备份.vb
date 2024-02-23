
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    
    Application.DisplayAlerts = False
    ans = MsgBox("Do you want to backup the file ?", vbYesNo)
    If ans = vbYes Then
        ThisWorkbook.SaveAs("D:\Document\01-Wanvog\03-Report\01-Inv. turns\ASInvTurnsBackup\AS INV. TURNS - " & Format(Now(), "yyyymmdd.hhmm") & ".xlsb")
    Else:
        Exit Sub
    End If
    Application.DisplayAlerts = True
End Sub





'D:\Document\06-Millennium\006-Project

Sub backupthisworkbook()
    'On Error Resume Next
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim wksSht As Worksheet
    Dim myfile As Object
    Dim fso As Object
    Dim myname As String
    Dim r As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")     ' fsoÓÃÓÚadd,edit,move,copy,delete ÎÄ¼þÓëÎÄ¼þ¼Ð
    Set myfile = CreateObject("Scripting.FileSystemObject")     ' fsoÓÃÓÚadd,edit,move,copy,delete ÎÄ¼þÓëÎÄ¼þ¼Ð
    
    
    myname = Dir(ThisWorkbook.Path & "\", vbDirectory)
    Do While myname <> ""
        If myname = "." Or myname = ".." Then
             Debug.Print myname
        ElseIf myname = "backup" Then
            GoTo 100
        Else
             Debug.Print myname
        End If
        myname = Dir
    Loop
    myfile.CreateFolder (ThisWorkbook.Path & "\backup")      ' ´´½¨backupÎÄ¼þ¼Ð
 
100

    
    Debug.Print ThisWorkbook.Path & "\" & ThisWorkbook.Name
    
    'a. ¸´ÖÆÎÄ¼þ µ½Ò»¸öÄ¿Â¼ÏÂ£º
    fso.CopyFile ThisWorkbook.Path & "\" & ThisWorkbook.Name, ThisWorkbook.Path & "\backup\"
    
    'b ¸ü¸ÄÎÄ¼þ¸´ÖÆºó µÄÄ¿Â¼ÏÂ ÎÄ¼þÃû£º
    'Name Â·¾¶ÏÂµÄÎÄ¼þÃû£¬±»¸ÄÎª AS ºóµÄÎÄ¼þÃû--ÕâÀïÒª´øÉÏnameµÄÂ·¾¶¼°¸ü¸ÄºóµÄÎÄ¼þÃû
    Name ThisWorkbook.Path & "\backup\" & ThisWorkbook.Name _
        As ThisWorkbook.Path & "\backup\MIL_CTN_Tracking - " & Format(Now(), "yyyy.mm.dd.hhmm") & ".xlsb"

    Set wksSht = Nothing

Application.ScreenUpdating = True
Application.DisplayAlerts = True

'MsgBox "Saved sheet 'InvTurnOverRatio' to D:\Document\03-Wanek3\00-Report\1.Inv. turns\BI\BW_MPT_INV_TURNOVER_RATIO.xlsx  "

End Sub




