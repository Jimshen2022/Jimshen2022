Sub Mail_small_Text_And_JPG_Range_Outlook()
    'Ron de Bruin, 12-03-2022
    'This macro use the function named : CopyRangeToJPG
    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String
    Dim MakeJPG As String

    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    strbody = "Dear Customer" & "<br><br>" & _
        "Below you find a picture of your data." & "<br>" & _
        "If you need more information let me know." & "<br><br>" & _
        "Regards Ron<br>"
              
    'Create JPG file of the range
    'Only enter the Sheet name and the range address
    MakeJPG = CopyRangeToJPG("Sheet3", "A1:Q16")

    If MakeJPG = "" Then
        MsgBox "Something go wrong, we can't create the mail"
        With Application
            .EnableEvents = True
            .ScreenUpdating = True
        End With
        Exit Sub
    End If

    On Error Resume Next
    With OutMail
        .To = "jishen@wanvogfurniture.com"
        .CC = ""
        .BCC = ""
        .Subject = "This is the Subject line"
        .Attachments.Add MakeJPG, 1, 0
        'Note: Change the width and height as needed
        .HTMLBody = "<html><p>" & strbody & "</p><img src=""cid:NamePicture.jpg"" width=750 height=700></html>"
        .Display 'or use .Send
    End With
    On Error GoTo 0

    Kill MakeJPG

    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub



Function CopyRangeToJPG(NameWorksheet As String, RangeAddress As String) As String
    'Ron de Bruin, 25-10-2019
    Dim PictureRange As Range

    With ActiveWorkbook
        On Error Resume Next
        .Worksheets(NameWorksheet).Activate
        Set PictureRange = .Worksheets(NameWorksheet).Range(RangeAddress)
        
        If PictureRange Is Nothing Then
            MsgBox "Sorry this is not a correct range"
            On Error GoTo 0
            Exit Function
        End If
        
        PictureRange.CopyPicture
        With .Worksheets(NameWorksheet).ChartObjects.Add(PictureRange.Left, PictureRange.Top, PictureRange.Width, PictureRange.Height)
            .Activate
            .Chart.Paste
            .Chart.Export Environ$("temp") & Application.PathSeparator & "NamePicture.jpg", "JPG"
        End With
        .Worksheets(NameWorksheet).ChartObjects(.Worksheets(NameWorksheet).ChartObjects.Count).Delete
    End With
    
    CopyRangeToJPG = Environ$("temp") & Application.PathSeparator & "NamePicture.jpg"
    Set PictureRange = Nothing
End Function

