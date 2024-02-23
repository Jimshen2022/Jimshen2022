'范例5.1 查找特定内容的单元格(查多个数值)
Sub FindNextCell()
    Dim StrFind As String
    Dim rng As range
    Dim FindAddress As String
    StrFind = InputBox("Please enter what you want to find: ")
    If Len(Trim(StrFind)) > 0 Then
        With Sheet1.range("a:a")
             .Interior.ColorIndex = 0
            Set rng =  .Find(What:=StrFind,  _
                    After:=.cells(.cells.count),  _
                    LookIn:=xlValues,  _
                    LookAt:=xlWhole,  _
                    SearchOrder:=xlByRows,  _
                    SearchDirection:=xlNext,  _
                    MatchCase:=False)
            If Not rng Is Nothing Then
                FindAddress = rng.address
                Do
                    rng.Interior.ColorIndex = 6
                    Set rng =  .Findnext(rng)
                Loop While Not rng Is Nothing _
                         And rng.address <> FindAddress
            End If
        End With
    End If
    Set rng = Nothing
End Sub

'范例5.2 使用Like进行模式匹配查找
Sub RngLike()
    Dim rng As range
    Dim r As Integer
    r = 1
    sheet1.range("a:a").clearcontents
    For Each rng In sheet2.range("a1:a40")
        If rng.text like "*a*" Then
            cells(r, 1) = rng.text
            r = r + 1
        End If
    Next
    Set rng = Nothing
End Sub

'范例5.2 使用Like进行模式匹配查找
Sub RngLike()
    Dim rng As Range
    Dim r As Integer
    r = 1
    Sheet2.Range("a:a").ClearContents
    Sheet2.Activate
    
    For Each rng In Sheet3.Range("a1:a50")
        If rng.Text Like "*[A,B,C,D]*" Then
            Cells(r, 1) = rng.Text
            r = r + 1
        End If
    Next
    Set rng = Nothing
End Sub
'范例5.2 使用Like进行模式匹配查找 *
Sub test()
    Dim i As Integer
    For i = 1 To 100
        If Range("a" & i) Like "VBA*" Then
            Range("a" & i).Interior.Color = 65535
        End If
    Next
End Sub

'范例5.2 使用Like进行模式匹配查找 ?
Sub test1()
    Dim i As Integer
    For i = 1 To 100
        If Range("a" & i) Like "V???B??" Then
            Range("a" & i).Interior.Color = 65535
        End If
    Next
End Sub

'范例5.2 使用Like进行模式匹配查找 方括号[]的使用
Sub test2()
    Dim i As Integer
    For i = 1 To 100
        If Range("a" & i) Like "[A-H]*" Then
            Range("a" & i).Interior.Color = 65535
        End If
    Next
End Sub
'范例5.2 使用Like进行模式匹配查找  井号(#)的使用
Sub test3()
    Dim i As Integer
    For i = 1 To 100
        If Range("a" & i) Like "##??????" Then
            Range("a" & i).Interior.Color = 65535
        End If
    Next
End Sub
'范例5.2 使用Like进行模式匹配查找  逻辑非(!)的使用

Sub test4()
    Dim i As Integer
    For i = 1 To 100
        If Range("a" & i) Like "#?[!0-9]*" Then
            Range("a" & i).Interior.Color = 65535
        End If
    Next
End Sub