Sub EachShtToWorkbook()
    Dim sht As Worksheet, strPath As String
    With Application.FileDialog(msoFileDialogFilePicker)
    '选择保存工作薄的文件路径
        If .Show Then strPath = .SelectedItems(1) Else Exit Sub
        '读取选择的文件路径,如果用户未选取文件路径则退出程序
    End With
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    Application.DisplayAlerts = False
    '取消显示系统警告消息，避免重名工作簿无法保存。当有重名工作簿时，会直接覆盖保存
    Application.ScreenUpdating = False
    For Each sht In Worksheets
        sht.Copy     '复制工作表，工作表单纯复制后，会成为活动工作表
        With ActiveWorkbook
            .SaveAs ActiveWorkbook.Path & sht.Name, xlWorkbookDefault
            '保存活动工作簿到指定路径下，以当前系统默认文件格式
            .Close True  '关闭工作簿并保存
        End With
    Next
    MsgBox "处理完成. ", , "提醒"
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
        

End Sub




Sub SaveSheets2()
    Dim wksSht As Worksheet
    Dim intCount As Integer
    Dim astrArr() As String
    For Each wksSht In Worksheets
        If VBA.InStr(1, wksSht.Name, "ÔÂ·Ý", vbTextCompare) > 0 Then
            intCount = intCount + 1
            ReDim Preserve astrArr(1 To intCount)
            astrArr(UBound(astrArr)) = wksSht.Name
        End If
    Next wksSht
    If intCount > 0 Then
        Worksheets(astrArr).Copy
        ActiveWorkbook.SaveAs Filename:= _
            ThisWorkbook.Path & "\SheetsBackup.xlsx"
    End If
    Set wksSht = Nothing
End Sub


Sub SaveSheets()
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim wksSht As Worksheet
    Dim intCount As Integer
    Dim astrArr() As String
    For Each wksSht In Worksheets
        If wksSht.Name = "STO" Then
           wksSht.Copy
           ActiveWorkbook.SaveAs Filename:= _
            ThisWorkbook.Path & "\BI\STO.xlsx"
            ActiveWorkbook.Close
        ElseIf wksSht.Name = "WANEK3_IN" Then
           wksSht.Copy
           ActiveWorkbook.SaveAs Filename:= _
            ThisWorkbook.Path & "\BI\WANEK3_IN.xlsx"
            ActiveWorkbook.Close
        ElseIf wksSht.Name = "Wanek3_DC_BW_OUT" Then
           wksSht.Copy
           ActiveWorkbook.SaveAs Filename:= _
            ThisWorkbook.Path & "\BI\Wanek3_DC_BW_OUT.xlsx"
            ActiveWorkbook.Close
        ElseIf wksSht.Name = "WANEK3_202_OUT" Then
           wksSht.Copy
           ActiveWorkbook.SaveAs Filename:= _
            ThisWorkbook.Path & "\BI\WANEK3_202_OUT.xlsx"
            ActiveWorkbook.Close
        ElseIf wksSht.Name = "BW_202_OUT" Then
           wksSht.Copy
           ActiveWorkbook.SaveAs Filename:= _
            ThisWorkbook.Path & "\BI\BW_202_OUT.xlsx"
            ActiveWorkbook.Close
        ElseIf wksSht.Name = "BW_202_OUT" Then
           wksSht.Copy
           ActiveWorkbook.SaveAs Filename:= _
            ThisWorkbook.Path & "\BI\BW_202_OUT.xlsx"
            ActiveWorkbook.Close
        ElseIf wksSht.Name = "DC_IN" Then
           wksSht.Copy
           ActiveWorkbook.SaveAs Filename:= _
            ThisWorkbook.Path & "\BI\DC_IN.xlsx"
            ActiveWorkbook.Close
        ElseIf wksSht.Name = "BW_IN" Then
           wksSht.Copy
           ActiveWorkbook.SaveAs Filename:= _
            ThisWorkbook.Path & "\BI\BW_IN.xlsx"
            ActiveWorkbook.Close
        ElseIf wksSht.Name = "UP" Then
           wksSht.Copy
           ActiveWorkbook.SaveAs Filename:= _
            ThisWorkbook.Path & "\BI\UP.xlsx"
            ActiveWorkbook.Close
        Else
            GoTo 100
            
        End If
100
    Next wksSht
    Set wksSht = Nothing
Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

