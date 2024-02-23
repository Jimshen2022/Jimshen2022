' ADO

Sub bySQL()
    Dim cnADO AS Object
    Dim rsADO AS Object
    Dim strSQL AS StrINg
    Dim i AS Long, strShtName, AShtName
    Set cnADO = CreateObject("ADODB.Connection")
    Set rsADO = CreateObject("ADODB.Recordset")
    cnADO.Open "Provider=Microsoft.ACE.OLEDB.12.0;" _
             & "Extended Properties=Excel 12.0;" _
             & "Data Source=" & ThisWorkbook.FullName
    AShtName = Split("一部门,二部门,三部门,四部门,后勤部", ",")
    For Each strShtName IN AShtName '多表合并语句
        strSQL = strSQL & "SELECT 姓名,工号 ,'" & strShtName & " ' AS 工作表名称 FROM [" & strShtName & "$]  UNION ALL "
    Next
    Set rsADO = cnADO.Execute(Left(strSQL, Len(strSQL) - 10))
    Cells.ClearContents
    For i = 0 To rsADO.Fields.Count - 1
        Cells(1, i + 1) = rsADO.Fields(i).Name
    Next i
    Range("a2").CopyFromRecordset rsADO
    rsADO.Close
    cnADO.Close
    Set cnADO = NothINg
    Set rsADO = NothINg
End Sub

'工作表批量转换为独立的工作簿
Sub EachShtToWorkbook()
    Dim sht AS Worksheet, strPath AS StrINg
    With Application.FileDialog(msoFileDialogFolderPicker)
   '选择保存工作薄的文件路径
        If .Show THEN strPath = .SELECTedItems(1) Else Exit Sub
        '读取选择的文件路径,如果用户未选取路径则退出程序
    End With
    If Right(strPath, 1) <> "\" THEN strPath = strPath & "\"
    Application.DisplayAlerts = False
    '取消显示系统警告和消息，避免重名工作簿无法保存。当有重名工作簿时，会直接覆盖保存。
    Application.ScreenUpdatINg = False '取消屏幕刷新
    For Each sht IN Worksheets '遍历工作表
        sht.Copy '复制工作表，工作表单纯复制后，会成为活动工作薄
        With ActiveWorkbook
            .SaveAS strPath & sht.Name, xlWorkbookDefault
            '保存活动工作薄到指定路径下，以当前系统默认文件格式
            .Close True '关闭工作薄并保存
        End With
    Next
    MsgBox "处理完成。", , "提醒"
    Application.ScreenUpdatINg = True '恢复屏幕刷新
    Application.DisplayAlerts = True '恢复显示系统警告和消息
End Sub


'如何删掉字符串中最后出现的连续数值？
Sub text()
    Dim arr, brr
    Dim i&, j&, n&, strTemp AS StrINg, strRes AS StrINg
    arr = Sheets("数据源").Range("A1").CurrentRegion
    For i = 2 To UBound(arr)
        b = False: bb = False
        strRes = "": s = arr(i, 1)
        For j = Len(s) To 1 Step -1
            strTemp = Mid(s, j, 1)
            If IsNumeric(strTemp) THEN
                If b = False THEN '如果还未找到过数值
                    b = True '标记开始找到数值
                Else '判断是不是第一个连续数值
                    If bb THEN strRes = strTemp & strRes  '第2个逻辑开关，表示连续数值已找完
                End If
            Else '如果不是数值
                If b = True THEN bb = True '标记找完连续数值了
                strRes = strTemp & strRes
            End If
        Next
        arr(i, 1) = strRes
    Next
    With Sheets("vba")
        .Cells.Clear
        .Range("A1").Resize(UBound(arr), 1) = arr
    End With
End Sub


'批量改变单元格部分字符格式

Sub MyCharacters()
    'ExcelHome技术论坛VBA编程学习与实践：看见星光
    Dim arr, s$, i&, l&, n&
    s = "领导" '需要改变格式的字符串
    n = Len(s) '变量s的长度
    arr = Range("a1:a" & Cells(Rows.Count, 1).End(xlUp).Row)
    For i = 1 To UBound(arr)
        l = INStr(1, arr(i, 1), s, vbTextCompare)
        '查找变量s在arr(i,1)中首次出现的位置，不区分字母大小写
        Do While l '如果l不为0，也就是存在s的话那么……
            With Cells(i, 1).Characters(l, n).Font
                 .Size = 15 '15号字体
                 .FontStyle = "加粗"
                 .Color = -16776961 '红色
            End With
            l = INStr(l + n, arr(i, 1), s, vbTextCompare)
            '寻找变量s下一个出现的位置
        Loop
    Next
    MsgBox "处理完毕!"
End Sub

'批量将图片从一张表格插入到另外一张表格
'一份工作簿有两张工作表。
'存放照片的工作表名为【照片】，需要插入图片的工作表名为【数据】。
'现在需要根据【数据】表的A列的图片名称，将【照片】表的照片批量插入到【数据】表的B列中去……
Sub INsertPicFromSheet()
    Dim shp AS Shape, rngData AS Range, rngPicName AS Range
    For Each shp IN ActiveSheet.Shapes
    '删除活动工作表原有照片
        If shp.Type = 13 THEN shp.Delete
    Next
    For Each rngData IN Range("a2", Cells(Rows.Count, 1).End(3))
        Set rngPicName = Sheets("照片").Cells.FINd(rngData.Value, , , xlWhole)
        '使用FINd方法在照片表的完整匹配姓名
        If Not rngPicName Is NothINg THEN rngPicName.Offset(0, 1).Copy rngData.Offset(0, 1)
        '如果有找到对应的姓名，则将照片复制粘贴到目标位置
    Next
End Sub

'如何使用SQL合并多工作表数据？
Sub SQL_UNION()
    Dim cnn AS Object, rst AS Object
    Dim strPath AS StrINg, str_cnn AS StrINg
    Dim strSQL AS StrINg, strTemp AS StrINg
    Dim sht AS Worksheet, strShtName AS StrINg
    Dim i AS Long
    Set cnn = CreateObject("adodb.connection")
    strPath = ThisWorkbook.FullName
    If Application.Version < 12 THEN
        str_cnn = "Provider=Microsoft.jet.OLEDB.4.0;Extended Properties=Excel 8.0;Data Source=" & strPath
    Else
        str_cnn = "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=Excel 12.0;Data Source=" & strPath
    End If
    cnn.Open str_cnn
    strTemp = "SELECT 姓名,语文,数学,英语,"
    For Each sht IN Worksheets
        strShtName = sht.Name
        If strShtName <> ActiveSheet.Name THEN
            strSQL = strSQL & strTemp & "'" & strShtName & "' AS 班级 FROM [" & strShtName & "$] UNION ALL "
        End If
    Next
    strSQL = Left(strSQL, Len(strSQL) - 11)
    Set rst = cnn.Execute(strSQL)
    Cells.ClearContents
    For i = 0 To rst.Fields.Count - 1
        Cells(1, i + 1) = rst.Fields(i).Name
    Next
    Range("a2").CopyFromRecordset rst
    cnn.Close
    Set cnn = NothINg
End Sub


'VBA 多表单数据汇总
Sub GatherData()
    
    Dim wb AS Workbook
    Dim sht AS Worksheet
    Dim oneSht AS Worksheet
    
    Set wb = ThisWorkbook
    Set sht = wb.Worksheets("汇总")
    
    With sht
        .UsedRange.Offset(1).ClearContents
        r = 2
    End With
    
    For Each oneSht IN wb.Worksheets
        If oneSht.Name <> sht.Name THEN
        
        With oneSht
            lAStRow = .Cells(Rows.Count, 1).End(xlUp).Row
            For i = 1 To lAStRow
                If .Cells(i, 1).Value = "序号" THEN
                    '送货单
                   kh = .Cells(i - 1, 3).Value
                   rq = .Cells(i - 1, 9).Value
                   dh = .Cells(i - 2, 9).Value
                   
                   For m = i + 1 To i + 5
                        If .Cells(m, 1).Value <> "" THEN
                                sht.Cells(r, 1) = kh
                                sht.Cells(r, 2) = rq
                                sht.Cells(r, 3) = dh
                                
                                sht.Cells(r, 4).Resize(1, 7).Value = .Cells(m, 2).Resize(1, 7).Value
                                
                                r = r + 1
                        End If
                   Next m
                   
                End If
            Next i
        End With
        
        
        End If
    Next oneSht
    
    Set wb = NothINg
    Set sht = NothINg
    Set oneSht = NothINg
    
End Sub

'批量取消工作表隐藏
Sub unShtVisible()
    Dim sht AS Worksheet
    For Each sht IN Worksheets '遍历工作表，设置可见
        sht.Visible = xlSheetVisible
    Next
End Sub

'如果只需要取消隐藏部分工作表，可以在代码中添加条件判断语句，将需要隐藏的工作表名称写在以下代码的第3行中，并以"/"作为分隔符合并即可。
Sub unShtVisible()
    Dim sht AS Worksheet, t
    t = "看见星光/Excel星球/Sheet5/" '将需要隐藏的工作表名称写在这
    For Each sht IN Worksheets '遍历工作表，设置可见
        If INStr(t, sht.Name &"/") THEN
            sht.Visible = xlSheetVisible
        End If
    Next
End Sub


'爬取TA在QQ空间的说说数据

Sub WebCrawlerQzone()
    Dim strURL AS StrINg
    Dim strCookie AS StrINg
    Dim strText AS StrINg
    Dim strGTK AS StrINg
    Dim strKey AS StrINg
    Dim strUserName AS StrINg
    Dim strMsg AS StrINg
    Dim INtPageNum AS Long
    Dim lngCreateTime AS Long
    Dim k AS Long
    Dim i AS Long
    Dim blnClick AS Boolean
    Dim objIE AS Object
    Dim objWINHTTP AS Object
    Dim objDIC AS Object
    Dim objDOM AS Object
    Dim objTagA AS Object
    Dim objList AS Object
    Dim objWINdow AS Object
    Dim vntTime AS Variant
    Dim vntQQNum AS Variant
    Set objDIC = CreateObject("scriptINg.dictionary")
    Set objIE = CreateObject("INternetExplorer.Application")
    Set objWINHTTP = CreateObject("WINHttp.WINHttpRequest.5.1")
    Set objDOM = CreateObject("htmlfile")
    Set objWINdow = objDOM.parentWINdow
    strURL = "https://xui.ptlogIN2.qq.com/cgi-bIN/xlogIN?"
    strURL = strURL & "proxy_url=https%3A//qzs.qq.com/"
    strURL = strURL & "qzone/v6/portal/proxy.html"
    strURL = strURL & "&appid=549000912"
    strURL = strURL & "&s_url=https%3A%2F%2Fqzs.qzone.qq.com" _
        & "%2Fqzone%2Fv5%2FlogINsucc.html%3Fpara%3Dizone"
    With objIE
        .navigate strURL
        .Visible = False
        vntTime = Timer
        Do While Timer < vntTime + 4
        Loop
        Do Until .readyState = 4
            DoEvents
        Loop
        For Each objTagA IN .document.getElementsByTagName("a")
            If objTagA.TabINdex = 2 THEN
                strUserName = objTagA.INnerText
                objTagA.Click
                blnClick = True
                Exit For
            End If
        Next
        If Not blnClick THEN
            MsgBox strUserName & "您的QQ软件未登录或QQ空间未开通。"
            Exit Sub
        End If
        vntTime = Timer
        Do While Timer < vntTime + 4
        Loop
        strCookie = .document.cookie
        .Quit
    End With
    strKey = Split(Split(strCookie, "p_skey=")(1), ";")(0)
    strGTK = strGetGTK(strKey)
    vntQQNum = [b1].Value
    strURL = "https://user.qzone.qq.com/"
    strURL = strURL & "proxy/domaIN/taotao.qq.com/"
    strURL = strURL & "cgi-bIN/emotion_cgi_msglist_v6?"
    strURL = strURL & "num=20"
    strURL = strURL & "&callback=_preloadCallback"
    strURL = strURL & "&format=jsonp"
    strURL = strURL & "&uIN=" & vntQQNum
    strURL = strURL & "&g_tk=" & strGTK
    ActiveSheet.UsedRange.Offset(2).ClearContents
    k = 3
    On Error Resume Next
    Application.ScreenUpdatINg = False
    Do While 1 = 1
        INtPageNum = INtPageNum + 20
        With objWINHTTP
            .Open "GET", strURL & "&pos=" & INtPageNum - 20, False
            .setRequestHeader "Cookie", strCookie
            .send
            strText = .responseText
        End With
        strText = Split(strText, "_preloadCallback(")(1)
        strText = Left(strText, INStrRev(strText, ")") - 1)
        objDOM.write "<script>var data=" & strText & "</script>"
        For i = 0 To objWINdow.eval("data.msglist.length") - 1
            k = k + 1
            Set objList = objWINdow.eval("data.msglist[" & i & "]")
            lngCreateTime = CallByName(objList, "created_time", VbGet)
            If Not objDIC.exists(lngCreateTime) THEN
                objDIC(lngCreateTime) = ""
            Else
                Exit Do
            End If
            Cells(k, 1) = CallByName(objList, "createTime", VbGet)
            Cells(k, 2) = CallByName(objList, "content", VbGet)
            Cells(k, 3) = CallByName(objList, "cmtnum", VbGet)
        Next i
    Loop
    [A3:C3] = Array("日期", "说说", "评论人数")
    Application.ScreenUpdatINg = True
    strMsg = "用户：" & strUserName & vbCrLf & "您好!"
    strMsg = strMsg & "目标QQ" & vntQQNum
    strMsg = strMsg & "的说说数据已抓取完成。"
    MsgBox strMsg
    Set objIE = NothINg
    Set objWINHTTP = NothINg
    Set objDOM = NothINg
    Set objWINdow = NothINg
    Set objDIC = NothINg
    Set objList = NothINg
End Sub
Function strGetGTK(ByVal strKey AS StrINg) AS StrINg
    Dim objNewDom AS Object
    Dim objNewWINdow AS Object
    Dim strJSON AS StrINg
    Set objNewDom = CreateObject("htmlfile")
    Set objNewWINdow = objNewDom.parentWINdow
    With objNewWINdow
        strJSON = "gtk=function(skey)"
        strJSON = strJSON & "{for(var hASh=5381,i=0,"
        strJSON = strJSON & "len=skey.length;i<len;++i)"
        strJSON = strJSON & "hASh+=(hASh<<5)"
        strJSON = strJSON & "+skey.charAt(i).charCodeAt();"
        strJSON = strJSON & "return hASh&2147483647}"
        strJSON = strJSON & "('" & strKey & "');"
        .execScript strJSON
        strGetGTK = .gtk
    End With
    Set objNewWINdow = NothINg
    Set objNewDom = NothINg
End Function

'INsertPicFromSheet2
Sub INsertPicFromSheet2()
    'ExcelHome VBA编程学习与实践 by:看见星光
    Dim rngData AS Range, rngWhere AS Range, cll AS Range
    Dim rngPicName AS Range, rngPic AS Range, rngPicPASte AS Range
    Dim shp AS Shape, sht AS Worksheet, bln AS Boolean
    Dim strWhere AS StrINg, strPicName AS StrINg, strPicShtName AS StrINg
    Dim x, y AS Long, lngYesCount AS Long, lngNoCount AS Long
    'On Error Resume Next
    Set rngData = Application.INputBox("请选择应插入图片名称的单元格区域", Type:=8)
    '用户选择需要插入图片的名称所在单元格范围
    Set rngData = INtersect(rngData.Parent.UsedRange, rngData)
    'INtersect语句避免用户选择整列单元格，造成无谓运算的情况
    If rngData Is NothINg THEN MsgBox "选择的单元格范围不存在数据！": Exit Sub
    strWhere = INputBox("请输入放置图片偏移的位置，例如上1、下1、左1、右1", , "右1")
    '用户输入图片相对单元格的偏移位置
    If Len(strWhere) = 0 THEN Exit Sub
    x = Left(strWhere, 1)
    '偏移的方向
    If INStr("上下左右", x) = 0 THEN MsgBox "你未输入偏移方位。": Exit Sub
    y = Val(Mid(strWhere, 2))
    '偏移的值
    SELECT CASE x
        CASE "上"
        Set rngWhere = rngData.Offset(-y, 0)
        CASE "下"
        Set rngWhere = rngData.Offset(y, 0)
        CASE "左"
        Set rngWhere = rngData.Offset(0, -y)
        CASE "右"
        Set rngWhere = rngData.Offset(0, y)
    End SELECT
    strPicShtName = INputBox("请输入存放图片的工作表名称", , "照片")
    For Each sht IN Worksheets
        If sht.Name = strPicShtName THEN bln = True
    Next
    If bln <> True THEN MsgBox "未找到保存图片的工作表：" & strPicShtName & vbCrLf & "程序退出。": Exit Sub
    Application.ScreenUpdatINg = False
    rngData.Parent.SELECT
    For Each shp IN ActiveSheet.Shapes
    '如果旧图片存放在目标图片存放范围则删除
        If Not INtersect(rngWhere, shp.TopLeftCell) Is NothINg THEN shp.Delete
    Next
    x = rngWhere.Row - rngData.Row
    y = rngWhere.Column - rngData.Column
    '偏移的纵横坐标
    For Each cll IN rngData
    '遍历选择区域的每一个单元格
        strPicName = cll.Text
        '图片名称
        If Len(strPicName) THEN
        '如果单元格存在值
            Set rngPicName = Sheets(strPicShtName).Cells.FINd(cll.Value, , , xlWhole)
            '使用FINd方法在照片表完整匹配姓名
            If Not rngPicName Is NothINg THEN
                Set rngPicPASte = cll.Offset(x, y)
                '粘贴图片的单元格
                Set rngPic = rngPicName.Offset(0, 1)
                '保存图片的单元格
                lngYesCount = lngYesCount + 1
                '累加找到结果的个数
                If lngYesCount = 1 THEN
                '设置放置图片单元格的行高和列宽，以适应图片的大小
                    rngPicPASte.RowHeight = rngPic.RowHeight
                    rngPicPASte.ColumnWidth = rngPic.ColumnWidth
                End If
                rngPicName.Offset(0, 1).Copy rngPicPASte
                '如果有找到对应的姓名，则将照片复制粘贴到目标位置
            Else
                lngNoCount = lngNoCount + 1
                '累加未找到结果的个数
            End If
        End If
    Next
    Application.ScreenUpdatINg = True
    MsgBox "共处理成功" & lngYesCount & "个对象，另有" & lngNoCount & "个非空单元格未找到对应的图片名称。"
End Sub



