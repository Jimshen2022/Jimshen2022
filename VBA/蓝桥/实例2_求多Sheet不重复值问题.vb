'问题的提出:
'一工作簿里面有3张工作表上，每张表格的A列都是姓名列，所有这些姓名中有些是重复的，要求编写一段代码，在另一个工作表上显示不重复的姓名。

Sub bcfz()
Dim i&, Myr&, Arr
Dim d, k, t, Sht As Worksheet
Set d = CreateObject("Scripting.Dictionary")
For Each Sht In Sheets
    If Sht.Name <> "Sheet4" Then
        Myr = Sht.[a65536].End(xlUp).Row
        Arr = Sht.Range("a2:a" & Myr)
        For i = 1 To UBound(Arr)
            d(Arr(i, 1)) = ""
        Next
    End If
Next
k = d.keys
Sheet4.[a3].Resize(d.Count, 1) = Application.Transpose(k)
Set d = Nothing
End Sub


'三?代码详解
'1、For Each Sht In Sheets ：For Each…Next循环结构，这种形式是VBA特有的，用于对对象的循环非常适用。意思是在所有的工作表中依次循环。
'2、If Sht.Name <> "Sheet4" Then ：如果这个工作表的名字不等于"Sheet4"时执行下面的代码。
'3、Myr = Sht.[a65536].End(xlUp).Row ：求得这个工作表A列有数据的最后一行的行数，把它赋给变量Myr。这里用了长整型数据类型(Long)，数据范围最大可到2,147,483,647，是为了避免数据很多的时候会超出整型数据类型(Integer)而出错，因为整型数据类型数据范围最大只到32,767。
'4 ?Arr = Sht.Range("a2:a" & Myr): 把A列数据赋给数组Arr?
'5、For i = 1 To UBound(Arr) ：For…Next循环结构，从1开始到数组的最大上限值之间循环。Ubound是VBA函数，返回数组的指定维数的最大值。
'6、d(Arr(i, 1)) = "" ：这句代码的意思就是把关键字Arr(i,1)加入字典，关键字对应的项为空，相当于字典中的这个关键字没有解释。和d.Add Arr(i,1), ""的效果相同，只是代码更简洁一些。
'7、k=d.keys ：把字典d中存在的所有的关键字赋给变量k。得到的是一个一维数组，下限为0，上限为d.Count-1。Keys是字典的方法，前面已经讲过了。
'8 ?Sheet4.[a3] .Resize(d.Count, 1) = Application.Transpose(k): 把字典d中所有的关键字赋给表4以a3单元格开始的单元格区域中?


