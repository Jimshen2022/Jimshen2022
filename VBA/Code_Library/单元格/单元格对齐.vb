




Sheet10.Range("e2:e21").Value = Date
    
    Sheet6.Select
    Sheet6.Range("A2") = "DataCollectedAt:  " & Format(Now, "HH:MM:SSam/pm,mmm.dd.yyyy")
    Sheet6.Range("A2").Font.color = -16776961
    Sheet8.Range("d:d, g:g,l:l").Font.color = -16776961  '不连续列设为红色
    Sheet8.Range("c:l").HorizontalAlignment = 3          '连续列设为置中对齐 3-置中对齐，1-靠左对齐
    Sheet25.Select




设置Excel中的一个或多个单元格甚至是一个区域的或者是被选中单元格的左对齐、友对齐、居中对齐、字体、字号、字型等属性。

　　①左对齐、右对齐、居中对齐

　　'选择区域或单元格右对齐　　
　　Selection.HorizontalAlignment = Excel.xlRight

　　'选择区域或单元格左对齐
　　Selection.HorizontalAlignment = Excel.xlLeft

　　'选择区域或单元格居中对齐　　
　　Selection.HorizontalAlignment = Excel.xlCenter

　　固定区域的对齐方式的代码：

　　Range("A1:A9").HorizontalAlignment = Excel.xlLeft

　　②字体、字号、字型

　　'当前单元格字体为粗体
　　Selection.Font.Bold = True

　　'当前单元格字体为斜体
　　Selection.Font.Italic = True

　　'当前单元格字体为宋体20号字

　　With Selection.Font
　　　.Name = "宋体"
　　　.Size = 20
　　End With
————————————————
版权声明：本文为CSDN博主「我在渤海之外」的原创文章，遵循CC 4.0 BY-SA版权协议，转载请附上原文出处链接及本声明。
原文链接：https://blog.csdn.net/u012867174/article/details/23272697


    