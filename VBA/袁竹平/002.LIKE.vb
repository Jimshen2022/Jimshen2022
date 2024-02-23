Like运算符的语法
Like运算符用于判断给定的字符串是否与指定的模式相匹配，其语法为：
结果=<字符串> Like <模式>
说明：
(1) <字符串>为文本字符串或者对包含文本字符串的单元格的引用，是要与<模式>相比较的字符串，数据类型为String型。
<模式>数据类型为String型，字符串中可以使用一些特殊字符，其它的字符都能与它们相匹配，其如下表1所示。
<模式>的字符                       与<字符串>匹配的文本
？                                          任意单个字符
*                                            零或者多个字符
#                                          任意单个数字(0-9)
[charlist]                                字符列表中的任意单个字符
[!charlist]                               不在字符列表中的任意单个字符
[ ]                                          空字符串(“ “)
(2) <结果>为Boolean型。如果字符串与指定的模式相匹配，则<结果>为True；否则<结果>为False。如果字符串或者模式Null，则结果为Null。
(3) Like运算符缺省的比较模式为二进制，因此区分大小写。可以用Option Compare语句来改变比较模式，如改变为文本比较模式，则不区分大小写。
(4) [Charlist]将模式中的一组字符与字符串中的一个字符进行匹配，可以包含任何一种字符，包括数字；在[Charlist]中使用连字号(-)产生一组字符来与字符串中的一个字符相匹配，如[A-D]与字符串相应位置的A、B、C或D匹配；在[Charlist]中可以产生多组字符，如[A-D H-J]；各组字符必须是按照排列顺序出现的；在Charlist的开头或结尾使用连字号(-)与连字号自身相匹配，例如[-H-N]与连字号(-)或H到N之间的任何字符相匹配。
在Charlist中的一个字符或者一组字符前加上！号，表明与该字符或该组字符之外的所有字符匹配，如[！H-N]与字符H-N范围之外的所有字符匹配；而在[]外使用！号则只匹配其自身。要使用任何特殊字符作为匹配字符，只需将它放在[]中即可，例如[？]表明要与一个问号进行匹配。
为了与左括号 ([)、问号 (?)、数字符号 (#) 和星号 (*) 等特殊字符进行匹配，可以将它们用方括号括起来。不能在一个组内使用右括号 (]) 与自身匹配，但在组外可以作为个别字符使用。

以实例来认识Like运算符
下面的代码演示了Like运算符在不同情况下所得的结果。
Sub testLikePattern()
Dim bLike1 As Boolean, bLike2 As Boolean
Dim bLike3 As Boolean, bLike4 As Boolean
Dim bLike5 As Boolean, bLike6 As Boolean
Dim bLike7 As Boolean
bLike1 = "aBBBa" Like "a*a"    ' 返回 True
bLike2 = "F" Like "[A-Z]"    ' 返回 True
bLike3 = "F" Like "[!A-Z]"    ' 返回 False
bLike4 = "a2a" Like "a#a"    ' 返回 True
bLike5 = "aM5b" Like "a[L-P]#[!c-e]"    ' 返回 True
bLike6 = "BAT123khg" Like "B?T*"    ' 返回 True
bLike7 = "CAT123khg" Like "B?T*"    ' 返回 False
MsgBox "Like运算符不同情形匹配结果:" & vbCrLf & " ""aBBBa"" Like ""a*a"" 结果为True." & _
     vbCrLf & """F"" Like ""[A-Z]""结果为True." & _
     vbCrLf & """F"" Like ""[!A-Z]""结果为False." & _
     vbCrLf & """a2a"" Like ""a#a""结果为True." & _
     vbCrLf & """aM5b"" Like ""a[L-P]#[!c-e]""结果为True." & _
     vbCrLf & """BAT123khg"" Like ""B?T*""结果为True." & _
     vbCrLf & """CAT123khg"" Like ""B?T*""结果为False."
End Sub

Like运算符的使用及示例

[示例一] 利用Like运算符自定义字符比较函数
IsLike函数非常简单，如果文本字符串与指定的模式匹配，该函数则返回True，代码如下：
Function IsLike(text As String,pattern As String) As Boolean
   IsLike= text Like pattern
End Function
该函数接受两个参数：
text：字符串或者是对包含字符串的单元格的引用。
pattern：包含有如上表1所示特殊字符的字符串。
函数的使用：在工作表中输入下面所示公式，可以查看函数的结果。
(1)下面的公式返回True。因为*匹配任意数量的字符。如果第一个参数是以“g”开始的任意文本，则返回True：
=IsLike(“guitar”,”g*”)
(2)下面的公式返回True。因为?匹配任意的单个字符。如果第一个参数是以“Unit12”，则返回False：
=IsLike(“Unit1”,”Unit?”)
(3)下面的公式返回True，原因是第一个参数是第二个参数的某个单个字符：
=IsLike(“a”,”[aeiou]”)
(4)如果单元格A1包含a,e,I,o,u,A,E,I,O或者U，那么下面的公式返回True。使用Upper函数作为参数，可以使得公式不区分大小写：
=IsLike(Upper(A1),Upper(“[aeiou]”))
(5)如果单元格A1包含以“1”开始并拥有3个数字的值(也就是100到199之间的任意整数)，那么下面的公式返回True：
=IsLike(A1,”1##”)

[示例二] 判断文本框的输入结果
打开一个工作簿，选择菜单“工具——宏——Visual Basic编辑器”或按Alt+F11组合键，打开VBE编辑器。在VBE编辑器中，选择菜单“插入——用户窗体”，新建一个用户窗体，点击控件工具箱中的“文本框”控件和“按钮”控件，在用户窗体上放置一个“文本框”和一个“按钮”，并在用户窗体中对它们的大小和位置进行合理调整，将按钮的标题改为

双击刚创建的“按钮”控件，并输入下面的代码：
Private Sub CommandButton1_Click()
Dim sEnd As String, sPattern As String
sEnd = "in Office"
sPattern = "[F W]*" & sEnd
If TextBox1.Text Like sPattern Then
    MsgBox "输入正确"
Else
    MsgBox "输入错误"
End If
End Sub
示例说明：本代码中[F W]*表示字符以F或W开头的字符串，使用&连接符将其与变量sEnd所代表的字符串“in Office”相连接。如果您在文本框中输入以字符F或字符W开头并以“in Office”结尾的句子，将显示“输入正确”消息框，否则将显示“输入错误”的消息框。在VBE编辑器中点击运行按钮或按F5键，运行该代码试试。当您在文本框中输入“Fanjy in Office”后，单击按钮，将显示“输入正确”消息框。注意，本示例区分大小写。

在Excel中实现字数统计
(1) 使用Len工作簿函数进行简单的字数统计
'① - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'对当前单元格进行字数统计
Sub TotalCellCharNum()
Dim i As Long
i = Len(ActiveCell.Value)
MsgBox "当前单元格的字数为：" & Chr(10) & i
End Sub
'② - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'对所选的单元格区域进行字数统计
Sub TotalSelectionCharNum()
Dim i As Long
Dim rng As Range
For Each rng In Selection
    i = i + Len(rng.Value)
Next rng
MsgBox "所选单元格区域的字数为：" & Chr(10) & i
End Sub

(2) 使用Like运算符进行较复杂的字数统计
'③ - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'对当前单元格中的文本分类进行字数统计
Sub SubTotalCellCharNum()
Dim str As String, ChineseChar As Long
Dim Alphabetic As Long, Number As Long
Dim i As Long, j As Long
j = Len(ActiveCell.Value)
For i = 1 To Len(ActiveCell)
    str = Mid(ActiveCell.Value, i, 1)
    If str Like "[一-龥]" = True Then
      ChineseChar = ChineseChar + 1 '汉字累加
    ElseIf str Like "[a-zA-Z]" = True Then
      Alphabetic = Alphabetic + 1 '字母累加
    ElseIf str Like "[0-9]" = True Then
      Number = Number + 1 '数字累加
    End If
Next
MsgBox "当前单元格中共有字数" & j & "个，其中：" & vbCrLf & "汉字：" & ChineseChar & "个" & _
     vbCrLf & "字母：" & Alphabetic & "个" & _
     vbCrLf & "数字：" & Number & "个", vbInformation, "文本分类统计"
End Sub
'④ - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'对所选的单元格区域中的文本分类进行字数统计
Sub SubTotalSelectionCharNum()
Dim str As String, ChineseChar As Long
Dim Alphabetic As Long, Number As Long
Dim i As Long, rng As Range, j As Long
For Each rng In Selection
    j = j + Len(rng.Value)
    For i = 1 To Len(rng)
      str = Mid(rng.Value, i, 1)
      If str Like "[一-龥]" = True Then
        ChineseChar = ChineseChar + 1 '汉字累加
      ElseIf str Like "[a-zA-Z]" = True Then
        Alphabetic = Alphabetic + 1 '字母累加
      ElseIf str Like "[0-9]" = True Then
        Number = Number + 1 '数字累加
      End If
    Next
Next
MsgBox "所选单元格区域中共有字数" & j & "个，其中：" & vbCrLf & "汉字：" & ChineseChar & "个" & _
     vbCrLf & "字母：" & Alphabetic & "个" & _
     vbCrLf & "数字：" & Number & "个", vbInformation, "文本分类统计"
End Sub

小结
上面对Like运算符的相关知识进行了介绍，在编写代码的过程中，可以灵活使用Like运算符。例如，可以用来进行字符串的比较，然后进行相应的统计或者执行进一步的操作，或者用来判断用户输入是否正确，以判断程序是否向下运行，等等