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
            .ChartType = xl3DColumnClustered
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

