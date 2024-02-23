' Wanek3 inventory turnover ratio 
'D:\Document\03-Wanek3\00-Report\1.Inv. turns\BW_DC_WN3 Inventory Turns Version01.xlsb

Sub BackupRawData()
    On Error Resume Next
    Application.ScreenUpdating = False
    Dim MyFile As Object
    Dim MyFiles As Object
    Dim SFilePath As String
    Dim DFilePath As String
    Dim fso As Object
    Dim NewFold As String
    
    NewFold = Str(Format(Date, "yyyymmdd")) & "\"    '新文件夹以当前日期命名
    SFilePath = "C:\Users\jishen\Downloads\"           'Source文件路径
    DFilePath = "D:\Document\03-Wanek3\00-Report\1.Inv. turns\DataBackup\" & NewFold   '目的文件路径
    
    Set MyFile = CreateObject("Scripting.FileSystemObject").Getfolder(SFilePath)   '取得某个文件夹下的所有文件
    Set fso = CreateObject("Scripting.FileSystemObject")     ' fso用于add,edit,move,copy,delete 文件与文件夹
    fso.CreateFolder (DFilePath)   '新建一个文件夹
    
    '遍历所有文件
    For Each MyFiles In MyFile.Files
        If MyFiles.Name Like "WANEK3*.xlsx" Then
            fso.CopyFile SFilePath & MyFiles.Name, DFilePath     '复制文件 --- 逗号前为source文件夹+文件名;  逗号后为复制到目的文件路径
            Name DFilePath & MyFiles.Name As DFilePath & MyFiles.Name & "-" & Format(Now(), "yyyymmdd.hhmm") & ".xlsx"
            '上一行第一个Name用于改文件名，或文件夹
        End If
    Next
    Application.ScreenUpdating = True
End Sub
