Sub InsertPicture2()
    Dim oPPT As Presentation
    Dim oSlide As Slide
    Dim nSlide As Byte
    Dim oCL As CustomLayout
    Dim Shp As Shape
    Dim myFile
    Dim filearr()
    Dim sPath As String
    Set oPPT = PowerPoint.ActivePresentation
    'sPath = PowerPoint.ActivePresentation.Path & "\"
    
    With Application.FileDialog(msoFileDialogFolderPicker)
         .Title = "Select Folder"
         .InitialFileName = "d:\"
        If  .Show Then
            sPath =  .SelectedItems(1)
        End If
    End With
    
    
    
    myFile = Dir(sPath & "\*.jpg")
    Do While myFile <> ""
        ReDim Preserve filearr(i)
        filearr(i) = Split(myFile, ".")(0)
        i = i + 1
        myFile = Dir
    Loop
    
    
    With oPPT
        For i = 0 To UBound(filearr) Step 6
            On Error Resume Next
            Set oCL =  .Slides(1).CustomLayout
            nSlide =  .Slides.Count
            If i Mod 6 = 0 And i <> 0 Then
                Set oSlide =  .Slides.AddSlide(1, oCL)
                Set Shp = oSlide.Shapes.AddPicture(sPath & "\" & filearr(i) & ".jpg", msoFalse, msoTrue, 80, 100, 204, 120)
                Set Shp = oSlide.Shapes.AddPicture(sPath & "\" & filearr(i + 1) & ".jpg", msoFalse, msoTrue, 300, 100, 204, 120)
                Set Shp = oSlide.Shapes.AddPicture(sPath & "\" & filearr(i + 2) & ".jpg", msoFalse, msoTrue, 80, 320, 204, 120)
                Set Shp = oSlide.Shapes.AddPicture(sPath & "\" & filearr(i + 3) & ".jpg", msoFalse, msoTrue, 300, 320, 204, 120)
                Set Shp = oSlide.Shapes.AddPicture(sPath & "\" & filearr(i + 4) & ".jpg", msoFalse, msoTrue, 300, 320, 204, 120)
                Set Shp = oSlide.Shapes.AddPicture(sPath & "\" & filearr(i + 5) & ".jpg", msoFalse, msoTrue, 300, 320, 204, 120)
                
                
            Else
                Set oSlide =  .Slides(nSlide)
                Set Shp = oSlide.Shapes.AddPicture(sPath & "\" & filearr(i) & ".jpg", msoFalse, msoTrue, 80, 100, 204, 120)
                Set Shp = oSlide.Shapes.AddPicture(sPath & "\" & filearr(i + 1) & ".jpg", msoFalse, msoTrue, 300, 100, 204, 120)
                Set Shp = oSlide.Shapes.AddPicture(sPath & "\" & filearr(i + 2) & ".jpg", msoFalse, msoTrue, 80, 320, 204, 120)
                Set Shp = oSlide.Shapes.AddPicture(sPath & "\" & filearr(i + 3) & ".jpg", msoFalse, msoTrue, 300, 320, 204, 120)
                Set Shp = oSlide.Shapes.AddPicture(sPath & "\" & filearr(i + 4) & ".jpg", msoFalse, msoTrue, 300, 320, 204, 120)
                Set Shp = oSlide.Shapes.AddPicture(sPath & "\" & filearr(i + 5) & ".jpg", msoFalse, msoTrue, 300, 320, 204, 120)
            End If
            
        Next
    End With
    MsgBox "完成"
End Sub


