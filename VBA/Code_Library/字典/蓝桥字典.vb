

'实例1  普通常见的求不重复值问题
'论坛网址：http://club.excelhome.net/thread-637004-1-1.html
'D:\Document\00-VBA\CodeLibrary\字典\蓝桥字典.xlsb

Sub cfz()

Application.ScreenUpdating = False
Dim i&, Myr&, Arr
Dim d, k, t
Set d = CreateObject("scripting.dictionary")
Myr = Sheet1.[a1048576].End(3).Row
Arr = Sheet1.Range("a1:g" & Myr)

For i = 2 To UBound(Arr)
    d(Arr(i, 3)) = d(Arr(i, 3)) + 1
Next

k = d.keys
t = d.items
Sheet2.Activate
[a2].Resize(d.Count, 1) = Application.Transpose(k)
[b2].Resize(d.Count, 1) = Application.Transpose(t)
[a1].Resize(1, 2) = Array("姓名", "重复个数")
Set d = Nothing
Erase Arr
Application.ScreenUpdating = True

End Sub



