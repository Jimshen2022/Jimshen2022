Sub Replacement()
    ' 范例6 替换单元格内字符串
    Sheet2.Range("a:a").Replace _
            What:="jim", Replacement:="Ashley",  _
            Lookat:=xlPart, searchorder:=xlByRows,  _
            MatchCase:=True ' MatchCase区分大小写,  Lookat:= xlWhole是指匹配全部搜索文本; xlPart指匹配任一部分搜索文本
End Sub



'5-1 使用Find方法查找特定信息
Sub FindCell()
    Dim strfind As String
    Dim rng As Range
    strfind = InputBox("请输入要查找的值：")
    If Len(Trim(strfind)) > 0 Then
        With Sheet1.Range("a:a")
            Set rng =  .Find(what:=strfind,  _
                    after:=.Cells(.Cells.Count),  _
                    LookIn:=xlValues,  _
                    lookat:=xlWhole,  _
                    searchorder:=xlByRows,  _
                    SearchDirection:=xlNext,  _
                    MatchCase:=False)
            If Not rng Is Nothing Then
                Application.GoTo rng, True
            Else
                MsgBox "没有找到匹配单元格!"
            End If
            Set rng = Nothing
        End With
    End If
    
End Sub


Sub FindNextCell()
    '查找多个值并用颜色标注
    Dim StrFind As String
    Dim rng As Range
    Dim FindAddress As String
    StrFind = InputBox("请输入要查找的值: ")
    If Len(Trim(StrFind)) > 0 Then
        With Sheet1.Range("a1").CurrentRegion
             .Interior.ColorIndex = 0
            Set rng =  .Find(what:=StrFind,  _
                    after:=.Cells(.Cells.Count),  _
                    LookIn:=xlValues,  _
                    lookat:=xlWhole,  _
                    searchorder:=xlByRows,  _
                    SearchDirection:=xlNext,  _
                    MatchCase:=False)
            If Not rng Is Nothing Then
                FindAddress = rng.Address
                Do
                    rng.Interior.ColorIndex = 6
                    Set rng =  .FindNext(rng)
                Loop While Not rng Is Nothing _
                         And rng.Address <> FindAddress
            End If
        End With
    End If
    Set rng = Nothing
    
End Sub

Sub RngLike()
    ' 5-2 使用like运算符进行模式匹配查找
    
    Dim rng As Range, r&
    r = 1
    Sheet1.Range("a:a").ClearContents
    For Each rng In Sheet2.Range("a1:a10000")
        If rng.Text Like "*a*" Then
            Sheet1.Cells(r, 1) = rng.Text
            r = r + 1
        End If
    Next
    Set rng = Nothing
End Sub







