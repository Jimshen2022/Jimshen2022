
Sub RegExpDemoSyntax()
    Dim objRegEx As Object
    Set objRegEx = CreateObject("vbscript.regexp")
    objRegEx.Pattern = "Name:(.*?),Phone:(\d+)"
    objRegEx.Global = True
    myString = "Name:张三丰,Phone:13801380000"
    Set objMH = objRegEx.Execute(myString)
    If objMH.Count > 0 Then
        With objMH(j)
            Debug.Print .SubMatches(0), .SubMatches(1)
        End With
    End If
    Set objRegEx = Nothing
End Sub


Sub RegExpDemoSyntax2()
    Dim 正则, 结果集合, 结果
    字符串 = Range("A2").Value
    '给对象指定正则表达式对象
    
    'CreateObject函数用于创建各种外部对象，对象的完整名称就是参数
    Set 正则 = CreateObject("vbscript.regexp")
    'Pattern后面写正则表达式
    正则.Pattern = "Name:(.*?),Phone:(\d+)"
    'Global值为True返回所有符合要求的结果，反之只返回第一个符合要求的结果
    正则.Global = True
    'Execute(字符串)
    Set 结果集合 = 正则.Execute(字符串)
    If 结果集合.Count > 0 Then
        i = 2
        For Each 结果 In 结果集合
            Range("B" & i) = 结果.SubMatches(0)
            Range("C" & i) = 结果.SubMatches(1)
            Range("D" & i) = 正则.Replace(字符串, "$1$2")
            i = i + 1
        Next
    End If
    Set 正则 = Nothing
    Set 结果集合 = Nothing
End Sub


Sub RegExpDemoSyntax3()
    Dim Regex, arr, result, Str
    Str = Range("A1").Value
    '给对象指定正则表达式对象
    
    'CreateObject函数用于创建各种外部对象，对象的完整名称就是参数
    Set Regex = CreateObject("vbscript.regexp")
    'Pattern后面写正则表达式
    Regex.Pattern = "Name:(.*?),Phone:(\d+)"
    'Global值为True返回所有符合要求的结果，反之只返回第一个符合要求的结果
    Regex.Global = True
    'Execute(Str)
    Set arr = Regex.Execute(Str)
    If arr.Count > 0 Then
        i = 2
        For Each result In arr
            Range("B" & i) = result.SubMatches(0)
            Range("C" & i) = result.SubMatches(1)
            Range("D" & i) = Regex.Replace(Str, "$1$2")
            i = i + 1
        Next
    End If
    Set Regex = Nothing
    Set arr = Nothing
End Sub





' 正则表达式示例1 提取字符串中的数字
Sub getNum1()
    ' 这种使用方式需要"工具""引用"
    ' 引用Microsoft VBScript Regular Expressions 5.5类库
    Dim reg As New RegExp
    With reg
        .Global = True
        .IgnoreCase = True
        .Pattern = "\d+"
    End With
    
   Dim mc As MatchCollection
   Dim m As Match
   Set mc = reg.Execute("123aaaaa987uiiui999")
   For Each m In mc
    MsgBox m.Value + 1
   Next
End Sub


' 正则表达式示例2 用"字符串"替换原字符串中符合匹配模式的部分
Sub getNum2()
    Dim arr
    arr = Split("A12B-R1E2W-E1T-R2T-Q1B2Y3U4D", "-") ' split(字符串,"分隔符")拆分字符串
    MsgBox "arr(0)=" & arr(0) & ";arr(1)=" & arr(1)
    MsgBox Join(arr, ",") ' join(数组,"分隔符")用分隔连接数组的每个元成一个字符串
    
    Dim arrStr() As String
    ReDim arrStr(LBound(arr) To UBound(arr)) ' 为动态数组分配存储空间
    With CreateObject("VBSCRIPT.REGEXP") ' 生成一个正则表达式对象实例
        For i = LBound(arr) To UBound(arr)
            .Global = True ' 设置全局可用，即替换所有符合匹配模式的字符串
            .Pattern = "[^A-Z]" ' 匹配模式为非大写字母
            arrStr(i) = .Replace(arr(i), "") ' 将arr(i)字符串中符合匹配模式的部分替换为空字符
        Next
    End With
    Cells.ClearContents
    Cells(1, 1).Resize(UBound(arr) + 1, 1) = Application.WorksheetFunction.Transpose(arrStr())
End Sub

