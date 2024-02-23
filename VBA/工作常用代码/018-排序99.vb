'jimshen tested for MIL STAGE REPORT -- UnKits on 2021/11/06
Sub Sorting() 'onhand

    Application.ScreenUpdating = False
    Sheet4.Select
    With Sheet4.Range("A1:K" & Range("A65536").End(xlUp).Row)
        .Sort Key1:=Range("K1"), Order1:=2, Header:=xlYes

        With Range("i1:K" & Range("A65536").End(xlUp).Row).Interior
        
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    .Columns("J:J").NumberFormat = "#,##0_);[Red](#,##0)"
    
    End With
    Application.ScreenUpdating = True
End Sub



'遇到多条件自定义排序——一个比较棘手的问题，请大神指教如何用VBA才能做到：
'A列（第一排序条件）：按一队、二队、三队
'B列（第二排序条件）：按甲1、甲2、甲3、甲4、甲5、甲6、甲7、甲8、甲9
'F列（第三排序条件）：按升序
'C列（第四排序条件）：按升序



Sub Macro1()
    Application.ScreenUpdating = False
    With Range("A3:R" & Range("A65536").End(xlUp).Row)
        .Sort Key1:=Range("F3"), Order1:=1, key2:=Range("C3"), order2:=1, Header:=xlNo
        Application.AddCustomList ListArray:=Array("甲1", "甲2", "甲3", "甲4", "甲5", "甲6", "甲7", "甲8", "甲9")
        .Sort Key1:=Range("b3"), Order1:=1, Header:=xlNo, OrderCustom:=Application.CustomListCount + 1
        Application.DeleteCustomList ListNum:=Application.CustomListCount
        Application.AddCustomList ListArray:=Array("一队", "二队", "三队")
        .Sort Key1:=Range("a3"), Order1:=1, Header:=xlNo, OrderCustom:=Application.CustomListCount + 1
        Application.DeleteCustomList ListNum:=Application.CustomListCount
    End With
    Application.ScreenUpdating = True
End Sub



'jimshen tested for MIL STAGE REPORT -- UnKits

Sub sorting2()

    Sheet4.Range("a1:k" & Range("a65563").End(xlUp).Row).Sort [k1], xlAscending, [a1], xlDscending, , , , xlYes '°´¿îºÅÓëÈÕÆÚÅÅÐò


End Sub


Sub FreeSort()
'eh技术论坛 VBA编程学习与实践 看见星光
Dim n&, rng As Range
Set rng = Range("e2:e" & Cells(Rows.Count, "e").End(xlUp).Row)
Application.AddCustomList (rng)
'增加一个自定义序列,该参数除了支持单元格对象，也支持数组。
n = Application.CustomListCount
'自定义序列的数目
Range("a:c").Sort key1:=[a1], order1:=xlAscending, Header:=xlYes, ordercustom:=n + 1
'使用自定义排序，ordercustom指定使用哪个自定义序列排序。
'当使用自定义排序时，需要将OrderCustom参数设置为指定的序列在自定义列表中的顺序加1
Application.DeleteCustomList n
'删除新增的自定义序列
End Sub










Sub soring_p&ic by Wanek
    ActiveWorkbook.Worksheets("UPHMO").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("UPHMO").Sort.SortFields.Add Key:=Range( _
        "H3:H" & LastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal                          'sort by FG_Due
    ActiveWorkbook.Worksheets("UPHMO").Sort.SortFields.Add Key:=Range( _
        "M3:M" & LastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal                           'sort by SHIFT
    ActiveWorkbook.Worksheets("UPHMO").Sort.SortFields.Add Key:=Range( _
        "L3:L" & LastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal                           'sort by LINE
    With ActiveWorkbook.Worksheets("UPHMO").Sort
        .SetRange Range("A2:P" & LastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With




Sub Qiliao() '
    
    Application.ScreenUpdating = False
    Dim arr1(), arr2()
    Dim a%, b%
    
    Sheets("Supply").Select
    Range("a1", Cells(Rows.Count, "f").End(xlUp)).Sort[c1], xlAscending, [f1], , xlAscending, [e1], xlAscending, xlYes '按款号与日期排序
    
    Sheet14.Range("a1", Cells(Rows.Count, "i").End(xlUp)).Sort[a1], xlAscending, , , , , , xlYes ''按款号排序 
End Sub



    Range("a1", Cells(Rows.Count, "k").End(xlUp)).Sort [k1], xlDscending, [a1], xlAscending, , , , , , xlYes '按款号与日期排序





Sub 排序之Sort()
    
    Sheet3.Range("A1:CN6").Sort Key1:=Range("N1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=n + 1, MatchCase:=True
    
    
    
    '语法 ：expression.Sort(Key1, Order1, Key2,  Type , Order2, Key3, Order3, Header, OrderCustom, MatchCase, Orientation, SortMethod, DataOption1, DataOption2, DataOption3)
    
    'expression 必须。一个表示 Range 对象的变量
    
    
    With Sheet3.Range("A1:CN6")
        
         .Sort Key1:=Sheet3.Range("N1") '..........第一排序关键字。
        
         .Sort Order1:=xlAscending '...............第一关键字排序方式xlAscending(或1)=升序，xlDescending(或2)=降序。
        
         .Sort Key2:=Range("F1") '.................第二关键字。
        
         .Sort Type :=xlChart '.....................指定要排序的元素。
        
         .Sort order2:=xlAscending '...............第二关键字排序方式xlAscending(或1)=升序，xlDescending(或2)=降序。
        
         .Sort key3:=Sheet3.Range("B1") '..........第三关键字。
        
         .Sort order3:=xlAscending '...............第三关键字排序方式xlAscending(或1)=升序，xlDescending(或2)=降序。
        
         .Sort Header:=xlGuess '...................指定第一行是否包含标题。xlGuess(或0)=工作表自己判断是否有标题，xlYes(或1)=强制第一行为列标题（不参与排序），xlNo(或2)=强制没有列标题（全部参与排序）
        
         .Sort OrderCustom:=n + 1 '................指定在自定义排序次序列表中的基于一的整数偏移(例：同一列中有ABCD，可以指定按DCBA或CDAB自定义排序，n变量可以是数组，也可以是单元格区域)
        
         .Sort MatchCase:=True '...................设置为True以执行区分大小写的排序, 设置为 False 以执行不区分大小写的排序; 否则为False 。不能用于数据透视表。
        
         .Sort Orientation:=xlSortColumns '........指定是应按行还是按列进行排序。 xlSortColumns(或1)按列排序。 xlSortRows(或2)=按行排序 。
        
         .Sort SortMethod:=xlPinYin '..............指定排序方法。xlPinYin(或1)=按字符的汉语拼音顺序排序，xlStroke()=按每个字符的笔划数排序。
        
         .Sort dataoption1:=xlSortNormal '.........指定如何对_Key1_中指定的范围内的文本进行排序;不适用于数据透视表排序。xlSortNormal(或0)=分别对数字和文本数据进行排序(默认值)，xlSortTextAsNumbers(或1)=将文本作为数字型数据进行排序。
        
         .Sort dataoption2:=xlSortNormal '.........指定如何对_Key2_中指定的范围内的文本进行排序;不适用于数据透视表排序。xlSortNormal(或0)=分别对数字和文本数据进行排序(默认值)，xlSortTextAsNumbers(或1)=将文本作为数字型数据进行排序。
        
         .Sort dataoption3:=xlSortNormal '.........指定如何对_Key3_中指定的范围内的文本进行排序;不适用于数据透视表排序。xlSortNormal(或0)=分别对数字和文本数据进行排序(默认值)，xlSortTextAsNumbers(或1)=将文本作为数字型数据进行排序。
        
    End With
    
    '******************************************* 关于第四个参数 Type 对应值的说明 *******************************************
    
    'xlChart........................(或-4109) = 图表
    'xlDialogSheet..................(或-4116) = 对话框工作表
    'xlExcel4IntlMacroSheet.........(或4)     = Excel 版本 4 国际宏工作表
    'xlExcel4MacroSheet.............(或3)     = Excel 版本 4 宏工作表
    'xlWorksheet....................(或-4167) = 工作表Worksheet
    
End Sub