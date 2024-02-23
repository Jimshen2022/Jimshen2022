'VBA实例-向未打开的工作簿中输入数据

'一：同目录下的员工花名册.xls中输入一个员工信息

'此段实例来源于《别怕，VBA其实很简单》一个实例

'4.6.4向未打开的工作簿中输入数据
'在当前文件所在目录中的"员工花名册.xlsx"工作簿中添加一条记录
'思路：是先打开工作簿---输入数据---再关闭工作簿

Sub WbInput()
Dim wb As String, xrow As Integer, arr                                             '定义变量 ,这里呢把wb直接设为了路径
wb = ThisWorkbook.Path & "\员工花名册.xlsx"                                         '将文件全名赋值给wb
Workbooks.Open (wb)                                                                '打开花名册文件
With ActiveWorkbook.Worksheets(1)                                                 '在工作簿中第一张表里添加记录
    xrow = .Range("a1").CurrentRegion.Rows.Count + 1                              '取得表格中第一条空行号
    arr = Array(xrow - 1, "刘伟", "刘伟", "刘伟", "刘伟", "刘伟")                  '将要录入的工作表的数据呢保存到数组arr中
    .Cells(xrow, 1).Resize(1, 6) = arr                                           '将数组写入单元格区域
End With
ActiveWorkbook.Close savechanges:=True                                          '关闭工作簿，并保存修改
End Sub



