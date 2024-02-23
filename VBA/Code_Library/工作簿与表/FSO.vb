
Private Sub Workbook_BeforeClose(Cancel As Boolean)

Application.ScreenUpdating = False
Application.DisplayAlerts = False
ans = MsgBox("Do you want to move the file to Ashton file server? ", vbYesNo)

    If ans = vbYes Then

        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")     ' fsoÓÃÓÚadd,edit,move,copy,delete ÎÄ¼þÓëÎÄ¼þ¼Ð
        fso.CopyFile "D:\Document\03-Wanek3\00-Report\0.Inv. Weekly report\Wanek Inv. weekly report - 2022.xlsb", "\\10.141.100.133\AshtonData\UPH FG Warehouse\Public\Inventory\Inventory-Jim\InventoryTurnsForSam\"
        fso.CopyFile "D:\Document\03-Wanek3\00-Report\0.Inv. Weekly report\Wanek Inv. weekly report - 2022.xlsb", "X:\"

    End If

Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
