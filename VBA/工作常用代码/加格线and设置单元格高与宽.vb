Sub formattextitem()
    
    '加一栏将item number 数字格式转为文本
    
    Dim i%, nrow&, arr()
    nrow = Sheet1.Range("a65563").End(3).Row
    arr = Sheet1.UsedRange
    Sheet1.Range("a2:k10000").ClearContents
    For i = 2 To nrow
        arr(i, 10) = Application.WorksheetFunction.Text(arr(i, 4), 0)
    Next
    
    With Sheet1.UsedRange
         .NumberFormat = "@"
        '.Columns("a:r").EntireColumn.AutoFit
         .Value = arr
        
    End With
    
    
    
    '给选中单元格区域加格线 20200925
    With Sheet1.Range("a1").CurrentRegion.Borders
         .LineStyle = xlContinuous
    End With
    
    
    With Sheet1.Range("a1").CurrentRegion
        
        '选中row 单元格内容置中
         .RowHeight = 45.75
         .VerticalAlignment = xlCenter
         .Orientation = 0
         .AddIndent = False
         .IndentLevel = 0
         .ShrinkToFit = False
         .ReadingOrder = xlContext
         .MergeCells = False
        
        '选中row 单元格内容向最左靠拢
         .HorizontalAlignment = xlLeft
         .VerticalAlignment = xlCenter
         .Orientation = 0
         .AddIndent = False
         .IndentLevel = 0
         .ShrinkToFit = False
         .ReadingOrder = xlContext
         .MergeCells = False
        
        Columns("a:h").EntireColumn.AutoFit
    End With
    
    
End Sub