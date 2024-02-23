Sub 自适应单元格()
    
    Dim i%, nrow&
    nrow = Sheet1.Range("a65563").End(3).Row
    
    For i = 2 To nrow
        
        With Sheet1.Range("i2:i" & nrow)
            
            '选中i列，自动缩小适应单元格
            
             .HorizontalAlignment = xlLeft
             .VerticalAlignment = xlCenter
             .WrapText = True
             .Orientation = 0
             .AddIndent = False
             .IndentLevel = 0
             .ShrinkToFit = False
             .ReadingOrder = xlContext
             .MergeCells = False
        End With
        
    Next
    
End Sub