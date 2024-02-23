'范例50 遍历工作表中的形状


Sub TraversalShapeOne()
    
    Dim i%
    For i = 74 To 79
        Sheet10.Shapes("TextBox" & " " & i).TextFrame.Characters.Text = "Jim"
    Next
End Sub


Sub TraversalShapeTwo()
    
    Dim i%, MyShape As Shape
    Dim MyCount%
    MyCount = 1
    
    For Each MyShape In Sheet10.Shapes
        If MyShape.Type = msoTextBox Then
            MyShape.TextFrame.Characters.Text = "BoxText" & MyCount & "!!!"
            MyCount = MyCount + 1
        End If
    Next
    Set MyShape = Nothing
    
End Sub