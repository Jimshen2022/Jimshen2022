'4.7.1 create a workbook
Sub WbAdd()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim wb As Workbook, sht As Worksheet
    Set wb = Workbooks.Add
    Set sht = wb.Worksheets(1)
    With sht
        .Name = "NameList"
        .Range("A1:F1") = Array("No.", "Name", "Male", "Birthday", "JoinTime", "Remark1")
    End With
    wb.SaveAs ThisWorkbook.Path & "\EmployeeList.xlsx"
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub


'4.7.2 Judge workbook is opening or not

Sub IsOpen()
    Dim i%
    For i = 1 To Workbooks.Count
        If Workbooks(i).Name = "EmployeeList.xlsx" Then
            MsgBox "the file is opening"
            Exit Sub
        End If
    Next
    MsgBox "The file is not opening"

End Sub

'Judge sht.name in worksheets, if not exist then add it and move to first
Sub ShtTest_1()
    Dim sht As Worksheet
    For Each sht In Worksheets
        If sht.Name = "NameList" Then
            sht.Move before:=Worksheets(1)
            Exit Sub
        End If
    Next
    Worksheets.Add(before:=Worksheets(1)).Name = "GradeOne"
End Sub

' or use below code
Sub ShtTest2()
    On Error Resume Next
    If Worksheets("GradeTwo") Is Nothing Then
        Worksheets.Add(before:=Worksheets(1)).Name = "GradeTwo"
    Else
        Worksheets("GradeOne").Move before:=Worksheets(1)
    End If
End Sub


'4.7.3 Judge workbook is exist or not

Sub TestFile()
    Dim fil As String
    fil = ThisWorkbook.Path & "\EmployeeList.xlsx"
    If Len(Dir(fil)) > 0 Then
        MsgBox "workbook is exist!"
    Else
        MsgBox "workbook is not exist!"
    End If

End Sub


'4.7.4 Entry Data into un-opening file
Sub testfile3()

    Dim wb As String, xrow As Integer, arr
    wb = ThisWorkbook.Path & "\EmployeeList.xlsx"
    Workbooks.Open (wb)
    With ActiveWorkbook.Worksheets(1)
        xrow = .Range("a1").CurrentRegion.Rows.Count  ' get the first empty row
        ' entry employee information into arr
        arr = Array(xrow, "ZhangJiao", "Female", #7/8/1987#, #9/1/2010#, "Year11 freshman")
        .Cells(xrow + 1, 1).Resize(1, 6) = arr
    End With
    ActiveWorkbook.Close savechanges:=True
    
End Sub

'4.7.5 隐藏活动工作表外的所有工作表


'4.7.6 批量新建工作表

'4.7.7 批量对数据分类

'4.7.10 summarized data from multiple workbooks under same file path 汇总同文件夹下多工作簿数据
Sub HzwWb()
	Dim r as Long, c as Long
	r = 1   '1 is table head rows
	c = 8   '8 is table head columns  
	Range(Cells(r + 1, "A"),Cells(65536,C)).ClearContents    '清除汇总表中原表数据
	Application.ScreenUpdating = False
	Dim FileName As String, wb as Workbook, sht as Worksheet, Erow As Long, Fn as String, arr as Variant
	Do While FileName <> ""
		If FileName <> ThisWorkbook.Name Then
			Erow = Range("A1").CurrentRegion.Rows.Count + 1
			fn = ThisWorkbook.Path & "\" & FileName
			Set wb = GetObject(fn)                       ' 将fn代表的工作簿对象赋给变量
			Set sht = wb.Worksheets(1)                   ' 汇总的是第1张工作表-- 内存中的
			'将数据表中的记录保存在arr数组里
			arr = sht.Range(sht.Cells(r+1,"A"), sht.Cells(65536,"B").End(3).Offset(0,8))
			'将数组arr中的数据写入工作表
			Cells(Erow,"A").Resize(Ubound(arr,1),Ubound(arr,2)) = arr
			wb.Close False
		End IF
		FileName = Dir    ' 自动取得上一次DIR目录下的下一个文件名
	Loop
	Application.ScreenUpdating = True
End Sub




'5.2 Worksheet事件

'5.2.2 worksheet_change 事件
Private Sub Worksheet_Change(ByVal Target As Range)
    Application.EnableEvents = False
    If Target.Column = 1 Then  '判断单元格是否为A列单元格
        MsgBox Target.Address & "  the cell was updated as : " & Target.Value '变量Target是程序运行的参数，代表工作表中被更改的单元格
    End If
    Application.EnableEvents = True
    
ThisWorkbook.Save

End Sub

'5.2.2 Worksheet_SelectionChange事件

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    MsgBox "当前选中的单元格区域为：" & Target.Address
    
End Sub


'光标一直在A列
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Column <> 1 Then
        Cells(Target.Row, "a").Select
    End If
    
End Sub

'Worksheet_Activate 事件： 自动提示工作表名
Private Sub Worksheet_Activate()
    MsgBox "当前活动工作表为： " & ActiveSheet.Name
End Sub

'Worksheet_Deactivate 事件： 禁止选中其他工作表
Private Sub Worksheet_Deactivate()
	msgbox "Not allow select other sheets except sheet1!" 
	worksheets("sheet1").Select

End sub



'5.2.3 Worksheet 事件列表

'5.3 WORKBOOK事件
'5.3.2 workbook open事件
Private Sub workbook_open()
	worksheeets(1).Select
End sub



'BeforeClose事件
Private Sub Workbook_BeforeClose(Cancel as Boolean)
	If msgbox("你确定要关闭工作簿? ", vbYesNo) = vbNo Then
		cancel = True   '取消关闭
End Sub



'WorkBook_SheetChange事件
'当工作簿里任意一个单元格被更改时,自动运行程序
Private Sub WorkBook_SheetChange(ByVal sht As Object, ByVal Target As Range)
    MsgBox "当前更改的工作表为: " & sht.Name & Chr(13) & _
            "发生更改的单元格地址为： " & Target.Address
End Sub


'5.4 别样的自动化
'5.4.1 MouseMove事件
Private Sub back_Click()
    cmd.Top = 15
    cmd.Left = 160
End Sub

Private Sub cmd_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim l As Integer, t As Integer
    l = Int(Rnd() * 10 + 125) * (Int(Rnd() * 3 + 1) - 2) ' Éú³ÉËæ»úÊý
    t = Int(Rnd() * 10 + 30) * (Int(Rnd() * 3 + 1) - 2) 'Éú³ÉËæ»úÊý
    cmd.Top = cmd.Top + t
    cmd.Left = cmd.Left + 1

End Sub



'5.4.2 不是事件的事件
'运行JIMS,再回到excel中  shift+E,自动执行test1

Sub JIMS()
    Application.OnKey "+e", "test1"
End Sub

Sub test1()
    MsgBox "hello world"
End Sub

'5.5.1 快速录入数据
Private Sub Worksheet_Change(ByVal Target As Range)
    '如果更改的单元各不是C列第3行以下的单元格或更改的单元格个数大于1时退出程序
    If Application.Intersect(Target, Range("c3:c65536")) Is Nothing Or Target.Count > 1 Then
        Exit Sub
    End If
    
    Dim i As Integer
    i = 3
    Do While Cells(i, "i").Value <> ""   ' 在参照表中循环
        '判断录入的字母与参照表的字母是否相符
        If UCase(Target.Value) = Cells(i, "I").Value Then
            Application.EnableEvents = False   '禁用事件,防止将字母改为商品名称时，再次执行该程序
                Target.Value = Cells(i, "I").Offset(0, 1).Value '写入产品信息
                Target.Offset(0, -1).Value = Date
                Target.Offset(0, 1).Value = Cells(i, "i").Offset(0, 2).Value   '写入商品代码
                Target.Offset(0, 2).Value = Cells(i, "i").Offset(0, 3).Value   '写入商品UP
                Target.Offset(0, 3).Select   '选中销售数量列，等待输入销售数量
            Application.EnableEvents = True
            Exit Sub
        End If
        i = i + 1
    Loop
End Sub




'5.5.2 监考哪一场
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Range("a2:t36").Interior.ColorIndex = xlNone   '清除单元格里原有底纹颜色
    '当选中的单元格个数大于1时，重新给Target赋值
    If Target.Count > 1 Then
        Set Target = Target.Cells(1)
    End If
    
    '当选中的单元格不包含指定区域的单元格时，退出程序
    If Application.Intersect(Target, Range("a2:t36")) Is Nothing Then
        Exit Sub
    End If
    
    Dim rng As Range
    For Each rng In Range("a2:t36")
        If rng.Value = Target.Value Then
            rng.Interior.ColorIndex = 39
        End If
        
    Next

End Sub



'5.5.3 让文件每隔一分钟自动保存一次

Private Sub Workbook_Open()
 Call otime    '打开工作簿后自动运行otime过程
End Sub

Sub otime()
    '一分钟自动运行Wbsave过程
    Application.OnTime Now() + TimeValue("00:00:30"), "wbsave"
End Sub

Sub wbsave()
    ThisWorkbook.Save
    Call otime
End Sub


'6.3 

























































