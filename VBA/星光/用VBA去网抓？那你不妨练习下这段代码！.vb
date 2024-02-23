Sub WebCrawlerDangD()
    Dim objXMLHTTP As Object
    Dim objDOM As Object
    Dim objDOMLi As Object
    Dim objShape As Shape
    Dim strURL As String
    Dim strText As String
    Dim strKey As String
    Dim strMsg As String
    Dim strMsgYesOrNo As String
    Dim strDOMLi As String
    Dim vntShapePic As Variant
    Dim intPageNum As Integer
    Dim intLiLength As Integer
    Dim lngaResult As Long
    Dim i As Long
    Dim k As Long
    strKey = [a2].Value
    If Len(strKey) = 0 Then
        MsgBox "未在A2单元格输入查询关键字。"
        Exit Sub
    End If
    ReDim aResult(1 To 7, 1 To 1)
    Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")
    Set objDOM = CreateObject("htmlfile")
    For intPageNum = 1 To 100
        strURL = "http://search.dangdang.com/?"
        strURL = strURL & "category_path=01.00.00.00.00.00#J_tab"
        strURL = strURL & "&act=input"
        strURL = strURL & "&key=" & strKey
        strURL = strURL & "&page_index=" & intPageNum
        With objXMLHTTP
             .Open "GET", strURL, False
             .send
            strText =  .responseText
        End With
        If InStr(strText, "没有找到") Then Exit For
        objDOM.body.innerHTML = strText
        Set objDOMLi = objDOM.getElementById("search_nature_rg").getElementsByTagName("li")
        intLiLength = objDOMLi.Length
        lngaResult = lngaResult + intLiLength
        ReDim Preserve aResult(1 To 7, 1 To lngaResult)
        For i = 0 To intLiLength - 1
            k = k + 1
            aResult(1, k) = k
            strDOMLi = objDOMLi(i).innerHTML
            strDOMLi = strDOMLi & "now_price>search_pre_price>search_discount>&nbsp;("
            aResult(4, k) = Val(Mid(Split(strDOMLi, "now_price>")(1), 2))
            aResult(5, k) = Val(Mid(Split(strDOMLi, "search_pre_price>")(1), 2))
            If aResult(5, k) = 0 Then aResult(5, k) = aResult(4, k)
            aResult(6, k) = Val(Split(strDOMLi, "search_discount>&nbsp;(")(1))
            If aResult(6, k) = 0 Then aResult(6, k) = ""
            With objDOMLi(i).getElementsByTagName("A")(0)
                aResult(3, k) =  .Title
                aResult(7, k) =  .href
            End With
            With objDOMLi(i).getElementsByTagName("IMG")(0)
                aResult(2, k) =  .src
                If Left(aResult(2, k), 4) <> "http" Then
                    aResult(2, k) =  .getAttribute("data-original")
                End If
            End With
        Next i
    Next intPageNum
    If k = 0 Then
        MsgBox "未找到符合条件的查询结果。"
        Exit Sub
    End If
    ActiveSheet.UsedRange.Offset(3).ClearContents
    Application.ScreenUpdating = False
    For Each objShape In ActiveSheet.Shapes
        If objShape. Type  = msoLinkedPicture Then objShape.Delete
    Next
    strMsg = "一共有" & k & "张图片需要导入Excel工作表。"
    If k > 200 Then strMsg = strMsg & "耗时过长！不建议导入！"
    strMsgYesOrNo = MsgBox("请选择是否需要导入图书图片！" _
             & VbCrLf & strMsg, vbYesNo)
    If strMsgYesOrNo = vbYes Then
        Const PIC_HEIGHT As Integer = 100
        Const RNG_HEIGHT As Integer = 110
        Const RNG_WIDTH As Byte = 16
        [B:B].ColumnWidth = RNG_WIDTH
        [A5].Resize(k, 1).EntireRow.RowHeight = RNG_HEIGHT
        For i = 1 To k
            Set vntShapePic = ActiveSheet.Pictures.Insert(aResult(2, i))
            With Cells(i + 4, 2)
                vntShapePic.Height = PIC_HEIGHT
                vntShapePic.Top = (RNG_HEIGHT - PIC_HEIGHT) / 2 +  .Top
                vntShapePic.Left = (.Width - vntShapePic.Width) / 2 +  .Left
            End With
            aResult(2, i) = ""
        Next i
    End If
    [a4:g4] = Array("序号", "封面", "书名", "现价", "定价", "折扣", "链接")
    [A5].Resize(k, UBound(aResult)) = Application.Transpose(aResult)
    Application.ScreenUpdating = True
    Set objXMLHTTP = Nothing
    Set objDOM = Nothing
    Set objDOMLi = Nothing
End Sub