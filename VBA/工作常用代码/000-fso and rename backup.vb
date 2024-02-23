'old 

Sub SaveSheetsForBI()
    'On Error Resume Next
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim wksSht As Worksheet

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")     ' fso用于add,edit,move,copy,delete 文件与文件夹
    
    For Each wksSht In Worksheets
        If wksSht.Name = "InvTurnOverRatio" Then
           'wksSht.Range("a9:e12").Cut Destination:=wksSht.Range("g1")
           wksSht.Copy
           wksSht.Columns("a:k").AutoFit
           'ActiveWorkbook.SaveAs filename:="X:\BW_MPT_INV_TURNOVER_RATIO.xlsx"
            'wksSht.Copy
           ActiveWorkbook.SaveAs filename:="D:\Document\03-Wanek3\00-Report\1.Inv. turns\BI\BW_MPT_INV_TURNOVER_RATIO.xlsx"
           ActiveWorkbook.Close
        End If
    Next wksSht
    
    fso.CopyFile "D:\Document\03-Wanek3\00-Report\1.Inv. turns\BI\BW_MPT_INV_TURNOVER_RATIO.xlsx", "D:\Document\03-Wanek3\00-Report\1.Inv. turns\BI\BwMptTurnOverBackup\"
    Name "D:\Document\03-Wanek3\00-Report\1.Inv. turns\BI\BwMptTurnOverBackup\BW_MPT_INV_TURNOVER_RATIO.xlsx" _
        As "D:\Document\03-Wanek3\00-Report\1.Inv. turns\BI\BwMptTurnOverBackup\" & "BW_MPT_INV_TURNOVER_RATIO-" & Format(Now(), "yyyy.mm.dd.hhmm") & ".xlsx"
    '上2行第一个Name用于改文件名，或文件夹
    Set wksSht = Nothing
Application.ScreenUpdating = True
Application.DisplayAlerts = True

MsgBox "Saved sheet 'InvTurnOverRatio' to D:\Document\03-Wanek3\00-Report\1.Inv. turns\BI\BW_MPT_INV_TURNOVER_RATIO.xlsx  "
End Sub



' 2022-9-10  D:\Document\06-Millennium\006-Project\2022 Sep Tracking

Sub backupthisworkbook()
    'On Error Resume Next
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim wksSht As Worksheet
    Dim myfile As Object
    Dim fso As Object
    Dim myname As String
    Dim r As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")     ' fso用于add,edit,move,copy,delete 文件与文件夹
    Set myfile = CreateObject("Scripting.FileSystemObject")     ' fso用于add,edit,move,copy,delete 文件与文件夹
    
    
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
    myfile.CreateFolder (ThisWorkbook.Path & "\backup")      ' 创建backup文件夹
 
100

    
    Debug.Print ThisWorkbook.Path & "\" & ThisWorkbook.Name
    
    'a. 复制文件 到一个目录下：
    fso.CopyFile ThisWorkbook.Path & "\" & ThisWorkbook.Name, ThisWorkbook.Path & "\backup\"
    
    'b 更改文件复制后 的目录下 文件名：
    'Name 路径下的文件名，被改为 AS 后的文件名--这里要带上name的路径及更改后的文件名
    Name ThisWorkbook.Path & "\backup\" & ThisWorkbook.Name _
        As ThisWorkbook.Path & "\backup\MIL_CTN_Tracking - " & Format(Now(), "yyyy.mm.dd.hhmm") & ".xlsb"

    Set wksSht = Nothing

Application.ScreenUpdating = True
Application.DisplayAlerts = True

'MsgBox "Saved sheet 'InvTurnOverRatio' to D:\Document\03-Wanek3\00-Report\1.Inv. turns\BI\BW_MPT_INV_TURNOVER_RATIO.xlsx  "

End Sub


