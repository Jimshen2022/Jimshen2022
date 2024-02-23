'范例52 自动插入图片
Sub InsertPicture()

'遍历某个文件夹下的所有jpg文件
    Dim MyFile As Object
    Dim MyFiles As Object
    Dim MyStr As String, i As Integer
    Set MyFile = CreateObject("Scripting.FileSystemObject").Getfolder("d:\Users\jishen\Pictures\B6 racking")
        '.Getfolder (ThisWorkbook.Path)
        K = 1
        With Sheet9
            For Each MyFiles In MyFile.Files
                If InStr(MyFiles.Name, ".jpg") <> 0 Then
                    K = K + 1
                    .Range("a" & K).Value = MyFiles.Name
                End If
            Next
        End With
    
    '在指定单元格插入图片，并且自适应
    Dim MyShape As Shape
    Dim r%, c%, picpath$, picrng As Range

    With Sheet9
        For Each MyShape In .Shapes
            If MyShape.Type = 13 Then
                MyShape.Delete
            End If
        Next
        For r = 2 To .Cells(.Rows.Count, 1).End(3).Row
            'For c = 1 To 8 Step 2
                picpath = "d:\Users\jishen\Pictures\B6 racking\" & .Cells(r, 1).Text
                If Dir(picpath) <> "" Then
                    Set MyShape = .Shapes.AddPicture(picpath, False, True, 56, 56, 56, 56)
                    Set picrng = .Cells(r, 2)
                    With MyShape
                        .LockAspectRatio = msoFalse
                        .Top = picrng.Top + 1
                        .Left = picrng.Left + 1
                        .Width = picrng.Width - 1.5
                        .Height = picrng.Height - 1.5
                        .TopLeftCell = ""
                    End With
                Else
                    .Cells(r, 2) = "暂无照片"
                End If
            'Next
        Next
        
    End With
    Set MyShape = Nothing
    Set picrng = Nothing
                
End Sub



'范例53 固定图片的尺寸和位置， 以下为单个图片，如果批量需要遍历

Sub FixedPicture()
    Dim picrng As Range
    Set picrng = Range("b3:b3")
    With Sheet9.Shapes(2)
     .LockAspectRatio = msoFalse
        .Rotation = 0
        .Top = picrng.Top - 1
        .Left = picrng.Left - 1
        .Width = picrng.Width + 1
        .Height = picrng.Height + 1
        
    End With
    Set picrng = Nothing
    
End Sub