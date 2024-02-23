Sub InsertPicture()
    Dim oPPT As Presentation
    Dim oSlide As Slide
    Dim nSlide As Byte
    Dim oCL As CustomLayout
    Dim Shp As Shape
    Dim myFile
    Dim filearr()
    Set oPPT = PowerPoint.ActivePresentation
    sPath = PowerPoint.ActivePresentation.Path & "\pic\"
    myFile = Dir(sPath & "*.png")
    Do While myFile <> ""
        ReDim Preserve filearr(i)
        filearr(i) = Split(myFile, ".")(0)
        i = i + 1
        myFile = Dir
    Loop
    
    
    With oPPT
        For i = 0 To UBound(filearr) Step 4
            On Error Resume Next
            Set oCL =  .Slides(1).CustomLayout
            nSlide =  .Slides.Count
            If i Mod 4 = 0 And i <> 0 Then
                Set oSlide =  .Slides.AddSlide(1, oCL)
                Set Shp = oSlide.Shapes.AddPicture(sPath & filearr(i) & ".png", msoFalse, msoTrue, 80, 100, 204, 120)
                Set Shp = oSlide.Shapes.AddPicture(sPath & filearr(i + 1) & ".png", msoFalse, msoTrue, 300, 100, 204, 120)
                Set Shp = oSlide.Shapes.AddPicture(sPath & filearr(i + 2) & ".png", msoFalse, msoTrue, 80, 320, 204, 120)
                Set Shp = oSlide.Shapes.AddPicture(sPath & filearr(i + 3) & ".png", msoFalse, msoTrue, 300, 320, 204, 120)
            Else
                Set oSlide =  .Slides(nSlide)
                Set Shp = oSlide.Shapes.AddPicture(sPath & filearr(i) & ".png", msoFalse, msoTrue, 80, 100, 204, 120)
                Set Shp = oSlide.Shapes.AddPicture(sPath & filearr(i + 1) & ".png", msoFalse, msoTrue, 300, 100, 204, 120)
                Set Shp = oSlide.Shapes.AddPicture(sPath & filearr(i + 2) & ".png", msoFalse, msoTrue, 80, 320, 204, 120)
                Set Shp = oSlide.Shapes.AddPicture(sPath & filearr(i + 3) & ".png", msoFalse, msoTrue, 300, 320, 204, 120)
            End If
            
        Next
    End With
    MsgBox "完成"
End Sub
