Sub sub_vlookup()
    Dim i As Long, j As Integer, dataArr, matchArr
    Dim dic As Object
    
    Set dic = CreateObject("Scripting.Dictionary")
    Sheets("数据源").Activate
    dataArr = [a1].CurrentRegion
    Sheets("匹配后的数据").Activate
    matchArr = [a1].CurrentRegion
    
    For i = 2 To UBound(dataArr)
        dic(dataArr(i, 1)) = dataArr(i, 2) & "," & dataArr(i, 3) & "," & dataArr(i, 4) & "," & dataArr(i, 5)
    Next i
    For i = 2 To UBound(matchArr)
        For j = 6 To 9
            If dic.exists(matchArr(i, 2)) Then
                matchArr(i, j) = Split(dic(matchArr(i, 2)), ",")(j - 6)
            End If
        Next j
    Next i
    Sheets("匹配后的数据").[a1].Resize(UBound(matchArr), UBound(matchArr, 2)) = matchArr
End Sub
