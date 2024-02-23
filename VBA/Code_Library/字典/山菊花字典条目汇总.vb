Sub tongji()
    Dim tp As Object, nRow%, Arr()
    
    Set tp = CreateObject("Scripting.dictionary") '创建一个字典对象
    
    '将数据保存到数组 Arr()中
    nRow = Range("a1").End(xlDown).Row
    Arr = Range("b2:c" & nRow).Value
    
    For i = 1 To nRow - 1 '循环读取数组
        For j = 1 To 2
            If Arr(i, j) <> "" Then
                tp(Arr(i, j)) = tp(Arr(i, j)) + 1 '以 姓名 为关键字，编辑条目。
                
                'MsgBox "候选人 " & Arr(i, j) & " 共得票数：" & tp(Arr(i, j)), 64, "计票"
                
            End If
        Next
    Next
    
    
    '输出结果到工作表
    Range("h2").Resize(tp.Count, 2).Value = WorksheetFunction.Transpose(Array(tp.keys, tp.items))
    
End Sub

'*****************        *****************        *****************
'1、编辑字典
'前面介绍过“创建字典”与“查字典”，这是字典的两个基本操作
'除此外
'还可以对字典进行编辑 (新华字典也在不断的修改，现在已经是第11版了)
'下面的代码就是对指定关键字的条目进行修改(加1)
'tp(Arr(i, j)) = tp(Arr(i, j)) + 1
'
'2、隐藏的 Add 方法
'上面的代码中没有使用 Add 方法，但从本地窗口及后面的语句中可以看出，字典中的条目被成功添加
'这是因为
'编辑字典条目tp(Arr(i, j))时，如果关键字Arr(i, j)在字典中不存在，它会自动添加该关键字
'就像唱票人唱到“唐僧”时，记票人发现黑板上没有“唐僧”的名字，他会先在候选人名字那行加上“唐僧”，然后在下面记上一横(先执行一次Add)
'
'3、Keys方法
'Keys方法返回一个数组，该数组包含一个 Dictionary 对象中的全部已有的关键字。
'返回的数组是一维数组，我们可以直接把它输出到工作表，如：
'    Range("l2:n2").Value = tp.keys
'也可以将所有关键字保存到一自定义数组，如：
'a= tp.keys
'
'4、Items 方法
'Items 方法返回一个包含 Dictionary 对象中所有条目的数组。
'用法参考“Keys方法”。
'
'
'
'*****************        *****************        *****************