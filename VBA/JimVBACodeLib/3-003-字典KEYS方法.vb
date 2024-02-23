'Keys方法
返回一个数组，其中包含了一个 Dictionary 对象中的全部现有的关键字。
object.Keys( )
其中 object 总是一个 Dictionary 对象的名称。

常用语句：
Dim d, k   
   Set d = CreateObject("Scripting.Dictionary")
   d.Add "a", "Athens"   
   d.Add "b", "Belgrade"
   d.Add "c", "Cairo"
   k=d.Keys
   [B1].Resize(d.Count,1)=Application.Transpose(k)
代码详解
1、Dim d, k ：声明变量，d见前例；k默认是可变型数据类型(Variant)。
2、k=d.Keys：把字典中存在的所有的关键字赋给变量k。得到的是一个一维数组，下限为0，上限为d.Count-1。这是数组的默认形式。
3、[B1].Resize(d.Count,1)=Application.Transpose(k) ：这句代码是很常用很经典的代码，所以这里要多说一些。
Resize是Range对象的一个属性，用于调整指定区域的大小，它有两个参数，第一个是行数，本例是d.Count，指的是字典中关键字的数量，整本字典中有多少个关键字，本例d.Count=3，因为有3个关键字。呵呵，是不是说多了。
第二个是列数，本例是1。这样＝左边的意思就是：把一个单元格B1调整为以B1开始的一列单元格区域，行数等于字典中关键字的数量d.Count，就是把单元格B1调整为单元格区域B1：B3了。
＝右边的k是个一维数组，是水平排列的，我们知道Excel工作表函数里面有个转置函数Transpose，用它可以把水平排列的置换成竖向排列。但是在VBA中不能直接使用该工作表函数，需要通过Application对象的WorksheetFunction属性来使用它。所以完整的写法是Application. WorksheetFunction.Transpose(k)，中间的WorksheetFunction可省略。现在可以解释这句代码了：把字典中所有的关键字赋给以B1单元格开始的单元格区域中。



'Items方法
返回一个数组，其中包含了一个 Dictionary 对象中的所有项目。
object.Items( )
其中 object 总是一个 Dictionary 对象的名称。

常用语句：
Dim d, t   
   Set d = CreateObject("Scripting.Dictionary")
   d.Add "a", "Athens"   
   d.Add "b", "Belgrade"
   d.Add "c", "Cairo"
   t=d.Items
   [C1].Resize(d.Count,1)=Application.Transpose(t)
代码详解
1、Dim d, t ：声明变量，d见前例；t默认是可变型数据类型(Variant)。
2、t=d.Items ：把字典中所有的关键字对应的项赋给变量t。得到的也是一个一维数组，下限为0，上限为d.Count-1。这是数组的默认形式。
3、[C1].Resize(d.Count,1)=Application.Transpose(t) ：有了上面Keys方法的解释这句代码就不用多说了，就是把字典中所有的关键字对应的项赋给以C1单元格开始的单元格区域中。




'Exists方法
如果 Dictionary 对象中存在所指定的关键字则返回 true，否则返回 false。
object.Exists(key)
参数
object 
必选项。总是一个 Dictionary 对象的名称。 
key 
必选项。需要在 Dictionary 对象中搜索的 key 值。

常用语句：
Dim d, msg$   
   Set d = CreateObject("Scripting.Dictionary")
   d.Add "a", "Athens"   
   d.Add "b", "Belgrade"
   d.Add "c", "Cairo"
   If d.Exists("c") Then
      msg = "指定的关键字已经存在。"
   Else
      msg = "指定的关键字不存在。"
   End If
代码详解
1、Dim d, msg$ ：声明变量，d见前例；msg$ 声明为字符串数据类型(String)，一般写法为Dim msg As String。String 的类型声明字符为美元号 ($)。
2、If d.Exists("c") Then：如果字典中存在关键字”c”，那么执行下面的语句。
3、msg = "指定的关键字已经存在。" ：把"指定的关键字已经存在。"字符串赋给变量msg。
4、Else ：否则执行下面的语句。
5、msg = "指定的关键字不存在。" ：把"指定的关键字不存在。"字符串赋给变量msg。
6、End If ：结束If …Else…Endif判断。


'Add方法
向 Dictionary 对象中添加一个关键字项目对。
object.Add (key, item)
参数
object 
必选项。总是一个 Dictionary 对象的名称。 
key 
必选项。与被添加的 item 相关联的 key。 
item 
必选项。与被添加的 key 相关联的 item。 
说明
如果 key 已经存在，那么将导致一个错误。

常用语句：
Dim d    
Set d = CreateObject("Scripting.Dictionary")
d.Add "a", "Athens"   
d.Add "b", "Belgrade"
d.Add "c", "Cairo"
代码详解
1、Dim d ：创建变量，也称为声明变量。变量d声明为可变型数据类型(Variant)，d后面没有写数据类型，默认就是可变型数据类型(Variant)。也有写成Dim d As Object的，声明为对象。
2、Set d = CreateObject("Scripting.Dictionary")：创建字典对象，并把字典对象赋给变量d。这是最常用的一句代码。所谓的“后期绑定”。用了这句代码就不用先引用c:\windows\system32\scrrun.dll了。
3、d.Add "a", "Athens"：添加一关键字”a”和对应于它的项”Athens”。 
4、d.Add "b", “Belgrade”：添加一关键字”b”和对应于它的项”Belgrade”。 
5、d.Add "c", “Cairo”：添加一关键字”c”和对应于它的项”Cairo”。




'Remove方法
Remove 方法从一个 Dictionary 对象中清除一个关键字，项目对。
object.Remove(key )
其中 object 总是一个 Dictionary 对象的名称。
key 
必选项。key 与要从 Dictionary 对象中删除的关键字，项目对相关联。 
说明
如果所指定的关键字，项目对不存在，那么将导致一个错误。

常用语句：
Dim d   
   Set d = CreateObject("Scripting.Dictionary")
   d.Add "a", "Athens"   
   d.Add "b", "Belgrade"
   d.Add "c", "Cairo"
   ……
   d.Remove(“b”)
代码详解
1、d.Remove(“b”)：清除字典中”b”关键字和与它对应的项。清除之后,现在字典里只有2个关键字了。



'RemoveAll方法
RemoveAll 方法从一个 Dictionary 对象中清除所有的关键字，项目对。
object.RemoveAll( )
其中 object 总是一个 Dictionary 对象的名称。
常用语句：
Dim d   
   Set d = CreateObject("Scripting.Dictionary")
   d.Add "a", "Athens"   
   d.Add "b", "Belgrade"
   d.Add "c", "Cairo"
   ……
   d.RemoveAll
代码详解
1、d.RemoveAll：清除字典中所有的数据。也就是清空这字典，然后可以添加新的关键字和项，形成一本新字典。

字典对象的属性有4个：Count属性、Key属性、Item属性、CompareMode属性。
Count属性
返回一个Dictionary 对象中的项目数。只读属性。
	object.Count
其中 object一个字典对象的名称。
常用语句：
Dim d,n%   
   Set d = CreateObject("Scripting.Dictionary")
   d.Add "a", "Athens"   
   d.Add "b", "Belgrade"
   d.Add "c", "Cairo"
   n = d.Count
代码详解
1、Dim d, n% ：声明变量，d见前例；n被声明为整型数据类型(Integer)。一般写法为Dim n As Integer 。 Integer 的类型声明字符为百分比号 (%)。
2、n = d.Count  ：把字典中所有的关键字的数量赋给变量n。本例得到的是3。



'Key属性
在 Dictionary 对象中设置一个 key。
object.Key(key) = newkey
参数：
object 
必选项。总是一个字典 (Dictionary) 对象的名称。 
key 
必选项。被改变的 key 值。 
newkey 
必选项。替换所指定的 key 的新值。 
说明
如果在改变一个 key 时没有发现该 key，那么将创建一个新的 key 并且其相关联的 item 被设置为空。
常用语句：
Dim d   
   Set d = CreateObject("Scripting.Dictionary")
   d.Add "a", "Athens"   
   d.Add "b", "Belgrade"
   d.Add "c", "Cairo"
   d.Key("c") = "d" 
代码详解
1、d.Key("c") = "d" ：用新的关键字”d”来替换指定的关键字”c”，这时，字典中就没有关键字c了，只有关键字d了，与d对应的项是”Cairo”。 




'Item属性
在一个 Dictionary 对象中设置或者返回所指定 key 的 item。对于集合则根据所指定的 key 返回一个 item。读/写。
object.Item(key)[ = newitem]
参数
object 
必选项。总是一个Dictionary 对象的名称。 
key 
必选项。与要被查找或添加的 item 相关联的 key。 
newitem 
可选项。仅适用于 Dictionary 对象；newitem 就是与所指定的 key 相关联的新值。 
说明
如果在改变一个 key 的时候没有找到该 item，那么将利用所指定的 newitem 创建一个新的 key。如果在试图返回一个已有项目的时候没有找到 key，那么将创建一个新的 key 且其相关的项目被设置为空。
常用语句：
Dim d   
   Set d = CreateObject("Scripting.Dictionary")
   d.Add "a", "Athens"   
   d.Add "b", "Belgrade"
   d.Add "c", "Cairo"
   MsgBox  d.Item("c") 
代码详解
1、d.Item("c") ：获取指定的关键字”c”对应的项。 
2、MsgBox   ：是一个VBA函数，用消息框显示。如果要详细了解MsgBox函数的，可参见我的另一篇文章“常用VBA函数精选合集”。http://club.excelhome.net/thread-387253-1-1.html

CompareMode属性
设置或者返回在 Dictionary 对象中进行字符串关键字比较时所使用的比较模式。
object.CompareMode[ = compare]
参数
object 
必选项。总是一个 Dictionary 对象的名称。 
compare 
可选项。如果提供了此项，compare 就是一个代表比较模式的值。可以使用的值是 0 (二进制)、1 (文本), 2 (数据库)。 
说明
如果试图改变一个已经包含有数据的 Dictionary 对象的比较模式，那么将导致一个错误。
常用语句：
Dim d   
   Set d = CreateObject("Scripting.Dictionary")
   d.CompareMode = vbTextCompare
   d.Add "a", "Athens"   
   d.Add "b", "Belgrade"
   d.Add "c", "Cairo"
   d.Add " B ", " Baltimore"
代码详解
1、d.CompareMode = vbTextCompare  ：设置字典的比较模式是文本，在这种比较模式下不区分关键字的大小写，即关键字”b”和”B”是一样的。vbTextCompare的值为1，所以上式也可写为 d.CompareMode =1 。如果设置为vbBinaryCompare（值为0），则执行二进制比较，即区分关键字的大小写，此种情况下关键字”b”和”B”被认为是不一样的。
2、d.Add " B ", " Baltimore" ：添加一关键字”B”和对应于它的项”Baltimore”。由于前面已经设置了比较模式为文本模式，不区分关键字的大小写，即关键字”b”和”B”是一样的，此时发生错误添加失败，因为字典中已经存在”b”了，字典中的关键字是唯一的，不能添加重复的关键字。






