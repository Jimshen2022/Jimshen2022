    
'范例27 自动建立工作表目录 ， 选中sheet13(3)时，自动将所有sheet名汇总过来
Private Sub worksheet_activate()
   Dim Sht As Worksheet
   Dim a As Integer
   Dim r As Integer
   r = Cells(Rows.Count, 1).End(3).Row
   a = 2
   If r > 1 Then Range("a2:a" & r).ClearContents
   For Each Sht In Worksheets
        If Sht.CodeName <> "Sheet1" Then
            Cells(a, 1).Value = Sht.Name
            a = a + 1
        End If
   Next
   Set Sht = Nothing

End Sub


'点击上述目录，自动跳转到对应Sheet

Private Sub worksheet_selectionchange(ByVal Target As Range)
    Dim r As Integer
    r = Cells(Rows.Count, 1).End(3).Row
    On Error Resume Next
    If Not Application.Intersect(Target, Range("a2:a" & r)) Is Nothing Then
        Sheets(Target.Text).Select
    End If
End Sub



'范例28 循环选择工作表
Sub ShtNext()     '选择最后一张工作表
    If ActiveSheet.Index < Worksheets.Count Then
        ActiveSheet.Next.Activate
    Else
        Worksheets(1).Activate
    End If
End Sub

Sub ShtPrevious()
'选择第一个工作表					
    If ActiveSheet.Index > 1 Then
        ActiveSheet.Previous.Activate
    Else
        Worksheets(Worksheets.Count).Activate
    End If
End Sub

