'一、问题的提出：
'有1、2、3…1000一千个数字，要求编写一段代码，在工作表的A列显示这些数被6除余1和余5的数字。


Sub 余1余5()  ‘by:狼版主
Dim dic As Object, i As Long, arr
Set dic = CreateObject("Scripting.Dictionary")
For i = 1 To 1000
dic.Add i & IIf(Abs(i Mod 6 - 3) = 2, "@", ""), ""
Next

arr = WorksheetFunction.Transpose(Filter(dic.keys, "@"))
[a1].Resize(UBound(arr), 1) = arr
[a:a].Replace "@", ""
Set dic = Nothing
End Sub


' 三、代码详解
' 1、Dim dic As Object, i As Long, arr  ：也可把字典变量dic声明为对象(Object)，i As Long是规范的写法，也可写成i& 。
' 2、dic.Add i & IIf(Abs(i Mod 6 - 3) = 2, "@", ""), "" ：这句代码的内容比较多，用了两个VBA函数IIf和Abs，用了一个Mod运算符。i Mod 6就是每一个数除6的余数，题目中有两个要求：余1和与5，为了从1到1000都同时能满足这两个要求，所以用了Abs(i Mod 6 - 3) = 2 ，Abs是取绝对值函数。另一个VBA函数IIf是根据判断条件返回结果，和If…Then判断结果类似；IIf(Abs(i Mod 6 - 3) = 2, "@", "") 这段的意思是如果符合判断条件，返回”@”否则返回空””。 i & IIf(Abs(i Mod 6 - 3) = 2, "@", "")的意思是把这个数与”@”或者”””连起来作为关键字加入字典dic，关键字相对应的项为空。比如当i=1时，1是满足上述表达式的，就把”1@” 作为关键字加入字典dic；当i=2时，2不满足上述表达式，就把”2” 作为关键字加入字典dic，关键字相对应的项都为空。
' 3、arr = WorksheetFunction.Transpose(Filter(dic.keys, "@")) ：这句代码的内容分为3部分，第1部分是Filter(dic.keys, "@") 其中的Filter是一个VBA函数，VBA函数就是可以直接在代码中使用的，我们平常使用的函数叫工作表函数，如Sum、Sumif、Transpose等等。Filter函数要求在一维数组中筛选出符合条件的另一个一维数组，式中的dic.keys正是一个一维数组。这里的筛选条件是”@”，也就是把字典关键字中含有@的关键字筛选出来组成一个新的一维数组，其下标从零开始。第2部分是用工作表函数Transpose转置这个新的一维数组，工作表函数的使用在前面keys方法一节已经说过了；第2部分是把转置以后的值赋给数组变量Arr。
' 呵呵，狼版主的代码是短了，我的解释却太长了。
' 4、[a1].Resize(UBound(arr), 1) = arr ：把数组Arr赋给[a1]单元格开始的区域中。
' 5、[a:a].Replace "@", ""  ：把A列中的所有的@都替换为空白，只剩下数字了。
