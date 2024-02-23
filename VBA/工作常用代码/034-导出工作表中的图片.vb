'范例48 导出工作表中的图片

Sub ExportPictures()
    Dim MyShape As Shape
    Dim FileName As String
    For Each MyShape In Sheet10.Shapes
       If MyShape.Type = msoPicture Then
         FileName = ThisWorkbook.Path & "\" & MyShape.Name & ".jpg"
         MyShape.Copy
         With Sheet10.ChartObjects.Add(0, 0, MyShape.Width, MyShape.Height).Chart
             .Paste
             .Export FileName
             .Parent.Delete
         End With
        End If
    Next
    Set MyShape = Nothing
End Sub