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


'范例49 在工作表中添加艺术字
Sub AddingWordArt()
    On Error Resume Next
    Sheet10.Shapes("MyShape").Delete
    Sheet10.Shapes.AddTextEffect _
    (PresetTextEffect:=msoTextEffect16, _
    Text:="Excel 2007", FontName:="宋体", _
    FontSize:=50, FontBold:=True, _
    FontItalic:=True, Left:=60, Top:=60).Name = "MyShape"
End Sub


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


'范例51 移动,旋转图形
Sub MoveAndRotate()
    Dim i%, j%
    With Sheet10.Shapes(1)
        For i = 1 To 3000 Step 5
            .Top = Sin(i * (3.1415926535 / 180)) * 100 + 100
            .Left = Cos(i * (3.1415926535 / 180)) * 100 + 100
            .Fill.ForeColor.RGB = i * 100
            For j = 1 To 20
                .IncrementRotation -2
                DoEvents
            Next
        Next
    End With
    
End Sub

'范例52 自动插入图片
Sub InsertPicture()

'遍历某个文件夹下的所有jpg文件
    Dim MyFile As Object
    Dim MyFiles As Object
    Dim MyStr As String, i As Integer
    Set MyFile = CreateObject("Scripting.FileSystemObject").Getfolder("d:\Users\jishen\Pictures\B6 racking")
        '.Getfolder (ThisWorkbook.Path)
        K = 1
        With Sheet9
            For Each MyFiles In MyFile.Files
                If InStr(MyFiles.Name, ".jpg") <> 0 Then
                    K = K + 1
                    .Range("a" & K).Value = MyFiles.Name
                End If
            Next
        End With
    
    '在指定单元格插入图片，并且自适应
    Dim MyShape As Shape
    Dim r%, c%, picpath$, picrng As Range

    With Sheet9
        For Each MyShape In .Shapes
            If MyShape.Type = 13 Then
                MyShape.Delete
            End If
        Next
        For r = 2 To .Cells(.Rows.Count, 1).End(3).Row
            'For c = 1 To 8 Step 2
                picpath = "d:\Users\jishen\Pictures\B6 racking\" & .Cells(r, 1).Text
                If Dir(picpath) <> "" Then
                    Set MyShape = .Shapes.AddPicture(picpath, False, True, 56, 56, 56, 56)
                    Set picrng = .Cells(r, 2)
                    With MyShape
                        .LockAspectRatio = msoFalse
                        .Top = picrng.Top + 1
                        .Left = picrng.Left + 1
                        .Width = picrng.Width - 1.5
                        .Height = picrng.Height - 1.5
                        .TopLeftCell = ""
                    End With
                Else
                    .Cells(r, 2) = "暂无照片"
                End If
            'Next
        Next
        
    End With
    Set MyShape = Nothing
    Set picrng = Nothing
                
End Sub


'范例53 固定图片的尺寸和位置

Sub FixedPicture()
    Dim picrng As Range
    Set picrng = Range("b3:b3")
    With Sheet9.Shapes(2)
     .LockAspectRatio = msoFalse
        .Rotation = 0
        .Top = picrng.Top - 1
        .Left = picrng.Left - 1
        .Width = picrng.Width + 1
        .Height = picrng.Height + 1
        
    End With
    Set picrng = Nothing
    
End Sub

'范例54 使用VBA自动生成图表
Sub ProdcutionCharts()
    Dim r As Integer
    Dim rng As Range
    Dim mychart As ChartObject
    On Error Resume Next
    With Sheet11
        .ChartObjects("mychart").Delete
        r = .Cells(.Rows.Count, 1).End(3).Row
        Set rng = .Range(.Cells(1, 1), .Cells(r, 2))
        Set mychart = .ChartObjects.Add(120, 40, 400, 250)
        mychart.Name = "mychart"
        With mychart.Chart
'            .ChartType = xl3DColumnClustered
'            .ChartType = xl3DArea
'            .ChartType = xl3DAreaStacked     '堆叠图
             .ChartType = BarStacked     '直方图

             
            .SetSourceData Source:=rng, PlotBy:=xlColumns
            .ApplyDataLabels ShowValue:=True
            .HasTitle = True
            With .ChartTitle
                .Text = "图表制作示例"
                .Font.Size = 14
            End With
        End With
    End With
    Set rng = Nothing
    Set mychart = Nothing
    

End Sub


'范例55 批量制作图表(根据姓名月份二维表制作直方图）
Sub ProductionCharts1()

    Dim mychart As ChartObject
    Dim i As Integer
    Dim r As Integer
    Dim m As Integer
    On Error Resume Next
    Sheet11.ChartObjects.Delete
    With Sheet11
        r = .Cells(.Rows.Count, 1).End(3).Row - 1
        m = Abs(Int(-(r / 4)))
        For i = 1 To r
            Set mychart = Sheet11.ChartObjects.Add _
                (Left:=(((i - 1) Mod m) + 1) * 350 - 340, _
                 Top:=((i - 1) \ m + 1) * 220 - 210, _
                 Width:=300, Height:=200)
            mychart.Name = .Range("a2").Offset(i - 1)
            With mychart.Chart
                .ChartType = xl3DColumnStacked
                .SetSourceData Source:=Sheet11.Range("b2:f2").Offset(i - 1), _
                    PlotBy:=xlRows
                .HasTitle = True
                .HasLegend = False
                With .ChartTitle
                    .Text = Sheet11.Range("a2").Offset(i - 1)
                    .Font.Name = "宋体"
                    .Font.Size = 12
                    
                End With
            End With
        Next
        
    End With
    Sheet11.Select
    Set mychart = Nothing

End Sub



'范例56 导出工作表中的图表--只导出一个图表，如果要导出所有图表，需要遍历
Sub exportchart()
    Dim chartpath As String
    chartpath = ThisWorkbook.Path & "\" & "mychart.jpg"
    On Error Resume Next
    Kill chartpath
        Sheet11.ChartObjects(1).Chart.Export FileName:=chartpath, filtername:="jpg"
        MsgBox "图表已保存在""" & ThisWorkbook.Path & """ 文件夹中！”"
        
End Sub

























