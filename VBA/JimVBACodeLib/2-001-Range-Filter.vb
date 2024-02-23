Sub Filter666()
    '取消筛选
    on error resume Next
    Dim i%, sht As Worksheet
    
    For Each sht In Worksheets
        ' 如果当前工作表为筛选模式，则取消
        If sht.AutoFilterMode = True Then sht.AutoFilterMode = 0
        ' 如果当前工作表没有筛选，则加上筛选
        If sht.AutoFilterMode = False Then sht.Range("a1").AutoFilter Field:=1
    Next
    
End Sub



Sub unfilter2() '先取消筛选，再加上筛选
    
    Dim i%, sht As Worksheet
    
    For Each sht In Worksheets
        ' 如果当前工作表为筛选模式，则取消
        If sht.AutoFilterMode = True Then sht.Range("a1:zz1").AutoFilter
        ' 如果当前工作表没有筛选，则加上筛选
        If sht.AutoFilterMode = False Then sht.Range("a1:zz1").AutoFilter Field:=1
    Next
    
End Sub


Sub unfilter1()
    '取消筛选
    ActiveSheet.AutoFilterMode = 0
    
End Sub

