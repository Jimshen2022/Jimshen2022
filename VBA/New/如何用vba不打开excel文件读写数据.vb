'如何用vba不打开excel文件读写数据？
'在编写vba代码的解决方案时，经常需要在不同的工作簿之间读写数据。
'接下来介绍几种在不同的excel工作簿之间读写数据的方法：
'一、打开读写法
'1、单个文件固定路径打开读写法：
'代码如下：
Sub QQ1722187970()
    Excel.Application.ScreenUpdating = False
    Excel.Application.DisplayAlerts = False
    Excel.Application.Calculation = xlCalculationManual
    Dim oWB As Workbook
    Dim oWK As Worksheet
    Dim sFilePath As String
    Dim iRow As Long
    '固定路径
    sFilePath = "E:\test.xlsx"
    Set oWB = Excel.Workbooks.Open(sFilePath)
    With oWB
        Set oWK = .Worksheets(1)
        With oWK
            iRow = .Range("a65536").End(xlUp).Row
            '***********************************
            '其它操作代码
            '***********************************
        End With
        Excel.Application.Calculation = xlCalculationAutomatic
        Excel.Application.DisplayAlerts = True
        Excel.Application.ScreenUpdating = True
        .Close
    End With
    MsgBox "操作完成!"
    Set oWK = Nothing
    Set oWB = Nothing
End Sub


'2、任意选择单个或多个文件打开读写法：
'代码如下:

Sub QQ1722187970()
    Excel.Application.ScreenUpdating = False
    Excel.Application.DisplayAlerts = False
    Excel.Application.Calculation = xlCalculationManual
    '选择路径读取打开法
    Dim oWB As Workbook
    Dim oWK As Worksheet
    Dim oFD As FileDialog
    Dim sFilePath As String
    Dim iRow As Long
    '创建一个选择文件对话框
    Set oFD = Excel.Application.FileDialog(msoFileDialogFilePicker)
    '声明一个变量用来存储选择的文件名
    Dim vrtSelectedItem As Variant
    With oFD
        '允许选择多个文件
        .AllowMultiSelect = True
        '使用Show方法显示对话框，如果单击了确定按钮则返回-1。
        If .Show = -1 Then
            '遍历所有选择的文件
            For Each vrtSelectedItem In .SelectedItems
                '获取所有选择的文件的完整路径,用于各种操作
                sFilePath = vrtSelectedItem
                Set oWB = Excel.Workbooks.Open(sFilePath)
                With oWB
                    Set oWK = .Worksheets(1)
                    With oWK
                        iRow = .Range("a65536").End(xlUp).Row
                        '***********************************
                        '其它操作代码
                        '***********************************
                    End With
                    Excel.Application.Calculation = xlCalculationAutomatic
                    .Close
                End With
            Next
        Set oWK = Nothing
        Set oWB = Nothing
        End If
    End With
    Excel.Application.DisplayAlerts = True
    Excel.Application.ScreenUpdating = True
End Sub


'3、任意选择文件夹及其子文件夹打开读写法：
'
'除了固定路径的单个文件和选择任意多个文件打开读写以外，我们往往还需要通过选择具体的文件夹，然后遍历文件夹内的所有文件进行打开读写，代码如下：
Sub QQ1722187970()
    Excel.Application.ScreenUpdating = False
    Excel.Application.DisplayAlerts = False
    Excel.Application.Calculation = xlCalculationManual
    Dim sPath As String
    '选择要操作的文件夹
    sPath = GetPath()
    If Len(sPath) Then
        '开始遍历选中的文件夹中的所有文件
        EnuAllFiles sPath, False
        MsgBox "操作完成!!!"
    End If
    Excel.Application.Calculation = xlCalculationAutomatic
    Excel.Application.DisplayAlerts = True
    Excel.Application.ScreenUpdating = True
End Sub
Sub EnuAllFiles(ByVal sPath As String, Optional bEnuSub As Boolean = False)
    '定义文件系统对象
    Dim oFso As Object
    Set oFso = CreateObject("Scripting.FileSystemObject")
    '定义文件夹对象
    Dim oFolder As Object
    Set oFolder = oFso.GetFolder(sPath)
    '定义文件对象
    Dim oFile As Object
    Dim oWB  As Workbook
    Dim oWK As Worksheet
    Dim oWB1  As Workbook
    Dim oWK1 As Worksheet
    Set oWB = Excel.ThisWorkbook
    Set oWK = oWB.Worksheets(1)
    iRow = oWK.Range("A65536").End(xlUp).Row
    '如果指定的文件夹含有文件
    If oFolder.Files.Count Then
        For Each oFile In oFolder.Files
            With oFile
                '输出文件所在的盘符
                Dim sDrive As String
                sDrive = .Drive
                '输出文件的类型
                Dim sType As String
                sType = .Type
                '输出含后缀名的文件名称
                Dim sName As String
                sName = .Name
                '输出含文件名的完整路径
                Dim sFilePath As String
                sFilePath = .Path
                '如果文件是Excel文件且不是隐藏文件
                If sType Like "*Excel*" And Not (sName Like "*~$*") Then
                    Set oWB1 = Excel.Workbooks.Open(sFilePath)
                    With oWB1
                        Set oWK1 = .Worksheets(1)
                        With oWK1
                            iRow = .Range("a65536").End(xlUp).Row
                            '***********************************
                            '其它操作代码
                            '***********************************
                        End With
                        Excel.Application.Calculation = xlCalculationAutomatic
                        .Close
                    End With
                Else

                End If
            End With
        Next
    '如果指定的文件夹不含有文件
    Else
    End If
    '如果要遍历子文件夹
    If bEnuSub = True Then
        '定义子文件夹集合对象
        Dim oSubFolders As Object
        Set oSubFolders = oFolder.SubFolders
        If oSubFolders.Count > 0 Then
            For Each oTempFolder In oSubFolders
                sTempPath = oTempFolder.Path
                Call EnuAllFiles(sTempPath, True)
            Next
        End If
        Set oSubFolders = Nothing
    End If
    Set oFile = Nothing
    Set oFolder = Nothing
    Set oFso = Nothing
End Sub
Function GetPath() As String
    '声明一个FileDialog对象变量
    Dim oFD As FileDialog
'    '创建一个选择文件对话框
'    Set oFD = Application.FileDialog(msoFileDialogFilePicker)
    '创建一个选择文件夹对话框
    Set oFD = Application.FileDialog(msoFileDialogFolderPicker)
    '声明一个变量用来存储选择的文件名或者文件夹名称
    Dim vrtSelectedItem As Variant
    With oFD
        '允许选择多个文件
        .AllowMultiSelect = True
        '使用Show方法显示对话框，如果单击了确定按钮则返回-1。
        If .Show = -1 Then
            '遍历所有选择的文件
            For Each vrtSelectedItem In .SelectedItems
                '获取所有选择的文件的完整路径,用于各种操作
                GetPath = vrtSelectedItem
            Next
            '如果单击了取消按钮则返回0
        Else
        End If
    End With
    '释放对象变量
    Set oFD = Nothing
End Function
Function GetFileName(ByVal sName As String)
    '获取不含后缀符的纯文件名的自定义函数
    Dim sTemp As String
    sTemp = sName
    '判断后缀名分隔符.的位置
    iPos = Len(sTemp) - VBA.InStr(1, VBA.StrReverse(sTemp), ".")
    If iPos <> 0 Then
        sTemp = Mid(sTemp, 1, iPos)
    End If
    '判断路径分隔符\的位置
    iPos = VBA.InStr(1, sTemp, "\")
    If iPos <> 0 Then
        '反转后好取字符
        iPos = VBA.InStr(1, VBA.StrReverse(sTemp), "\")
        sTemp = Mid(VBA.StrReverse(sTemp), 1, iPos - 1)
        sTemp = VBA.StrReverse(sTemp)
    End If
    GetFileName = sTemp
End Function


'4、总结

'以上介绍的三种方法基本涵盖了所有的在不同excel工作簿之间的读写数据的情况。

'以上介绍的三种方法在读写其它excel工作簿的数据时，本质上都是用Workbooks对象的Open方法先打开要读写的excel工作簿，然后再进行操作。

