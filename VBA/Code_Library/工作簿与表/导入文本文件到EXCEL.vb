'9-134 loading text files

'使用查询表导入
Sub AddQuery()
    With Sheet23
        .UsedRange.ClearContents
        With .QueryTables.Add(Connection:="text;" & "C:\Users\jishen\Downloads" & "\STO_report.txt", Destination:=.Range("a1"))
            .TextFileCommaDelimiter = False
            .Refresh
        End With
        .Select
        
    End With

End Sub
   

'使用Open语句导入

Sub openText()
    Dim MyText As String
    Dim MyArr() As String
    Dim c As Integer
    Dim r As Integer
    r = 1
    With Sheet23
        .UsedRange.ClearContents
        Open "C:\Users\jishen\Downloads\STO_report.txt" For Input As #1
        Do While Not EOF(1)
            Line Input #1, MyText
            MyArr = Split(MyText, ",")
            For c = 0 To UBound(MyArr)
                .Cells(r, c + 1) = MyArr(c)
            Next
            r = r + 1
        Loop
        Close #1
        .Select
    End With
        
End Sub



'使用OpenText方法导入

Sub OpenText2()
    Sheet23.UsedRange.ClearContents
    Workbooks.openText FileName:="C:\Users\jishen\Downloads\STO_report.txt", startrow:=1, DataType:=xlDelimited, comma:=False
    With ActiveWorkbook
        With .Sheets("STO_report").Range("a1").CurrentRegion
            ActiveWorkbook.Sheets("STO_report").Range("a1").Resize(.Rows.Count, .Columns.Count).Value = .Value
            'ActiveWorkbook.Sheets("STO_report").Copy ThisWorkbook.Sheets("Sheet2")
        End With
     .Close False
    End With
    Sheet23.Select
        
End Sub



















