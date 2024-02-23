Sub POAllocationForShortage() 'ready 12  allocation PO and MO for Trips
    
    '    Application.ScreenUpdating = False
    '    Application.Calculation = xlCalculationManual
    '    Application.StatusBar = "Calculate picked and packed RP Orders, please wait ......"
    
    Dim arr1(), arr2()
    Dim a%, b%
    
    Sheet10.Activate 'shortageSheet
    'Range("b1", Cells(Rows.Count, "g").End(xlUp)).Sort [d1], xlAscending, [b1], , xlAscending, , , xlYes  '°´¿îºÅÓëÈÕÆÚÅÅÐò
    arr1 = Range("a2", Cells(Cells(Rows.Count, "n").End(xlUp).Row, "o")) 'Êý×é¸³Öµ
    Sheet2.Activate ' supplySheet
    'Range("a1", Cells(Rows.Count, "c").End(xlUp)).Sort [a1], xlAscending, [c1], , xlAscending, , , xlYes   '°´¿îºÅÓëÈÕÆÚÅÅÐò
    arr2 = Range("a2", Cells(Rows.Count, "y").End(xlUp)) 'Êý×é¸³Öµ
    
    'For a = LBound(arr1) To UBound(arr1)
    '    arr1(a, 6) = arr1(a, 6) * 1   '½«ÎÄ×Ö¸ÄÎªÊýÖµ
    'Next
    
    Sheet10.Activate
    Range("o2:o66365").ClearContents
    
100
    For a = LBound(arr1) To UBound(arr1)
        For b = LBound(arr2) To UBound(arr2)
            If arr1(a, 5) = arr2(b, 2) And arr1(a, 14) <= arr2(b, 3) And arr1(a, 14) <> 0 Then '¶©µ¥¿îºÅÓëPO¿îºÅÏàÍ¬£¬ÇÒ ¶©µ¥ÐèÇó<=POÊýÁ¿ ÇÒ ¶©µ¥ÐèÇó²»Îª0
                arr2(b, 3) = arr2(b, 3) - arr1(a, 14) 'POÊýÁ¿ = POÊýÁ¿ - ¶©µ¥ÐèÇó
                arr1(a, 14) = arr1(a, 14) - arr1(a, 14) '¶©µ¥ÐèÇó  = ¶©µ¥ÐèÇó - ¶©µ¥ÐèÇó
                Sheet10.Cells(a + 1, "o") = arr2(b, 1) 'Ô¤¼ÆÆëÁÏÈÕ = POµÄÔ¤¼Æ½øÁÏÈÕÆÚ   ¡®cells(a+1)ÊÇÎªÁË¸úÊý×éµÚÒ»ÐÐ¶ÔÆë
                arr1(a, 15) = arr2(b, 1)
            ElseIf arr1(a, 5) = arr2(b, 2) And arr1(a, 14) > arr2(b, 3) And arr2(b, 3) <> 0 Then 'Òà»ò£¬¶©µ¥¿îºÅÓëPO¿îºÅÏàÍ¬£¬ÇÒ ¶©µ¥ÐèÇó>POÊýÁ¿ ÇÒ PO²»Îª0
                arr1(a, 14) = arr1(a, 14) - arr2(b, 3) '¶©µ¥È±»õÊýÁ¿ = ¶©µ¥È±»õÊýÁ¿ - POÊýÁ¿
                arr2(b, 3) = arr2(b, 3) - arr2(b, 3) 'POÊýÁ¿ = POÊýÁ¿ - POÊýÁ¿
                If arr1(a, 14) <> 0 Then 'Èç¹û¶©µ¥ÐèÇó<>0, »Øµ½100
                    GoTo 100
                End If
            End If
        Next
    Next
    
    
    
    Sheet10.Select
    
    Dim i%
    For i = 2 To [M66365].End(3).Row
        If Cells(i, "O") = "" Then Cells(i, "O") = "PO Uncovered"
    Next
    
    Columns("N:O").Select
    Range("O1").Activate
    With Selection
         .HorizontalAlignment = xlGeneral
         .VerticalAlignment = xlBottom
         .WrapText = False
         .Orientation = 0
         .AddIndent = False
         .IndentLevel = 0
         .ShrinkToFit = False
         .ReadingOrder = xlContext
         .MergeCells = False
    End With
    With Selection
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlBottom
         .WrapText = False
         .Orientation = 0
         .AddIndent = False
         .IndentLevel = 0
         .ShrinkToFit = False
         .ReadingOrder = xlContext
         .MergeCells = False
    End With
    Range("O11").Select
    
    
    '    Application.Calculation = xlCalculationAutomatic
    '    Application.ScreenUpdating = True
    '    Application.StatusBar = False
    
End Sub