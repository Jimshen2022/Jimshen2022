'第4章 Shape,Chart

'范例47 在工作表中添加图形

Sub AddingGraphics()
    Dim MyShape As Shape
    On Error Resume Next
    Sheet1.Shapes("MyShape").Delete
    Set MyShape = Sheet8.Shapes.AddShape(msoShapeRectangle, 40, 130, 280, 30)
    
    With MyShape
        .Name = "MyShape"
        With .TextFrame.Characters
            .Text = "click will select sheet2!"
            With .Font
                .Size = 20
                .ColorIndex = 5
            End With
        End With
        With .line
            .Weight = 1
            .Style = msoLineSingle
            .Transparency = 0.5            '透明度
            .ForeColor.SchemeColor = 40    '前景色
            .BackColor.RGB = RGB(255, 255, 255)
        End With
        
        With .Fill
            .Transparency = 0.5
            .ForeColor.SchemeColor = 41
            .OneColorGradient 1, 4, 0.23
        End With
        .Placement = xlFreeFloating
    End With
    Sheet8.Hyperlinks.Add anchor:=MyShape, Address:="", _
        SubAddress:="Sheet7!a1", ScreenTip:="选择Sheet7!"
    Set MyShape = Nothin
    
End Sub