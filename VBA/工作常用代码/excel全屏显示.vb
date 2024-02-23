Sub showfullscreen()
    'excel全屏显示
    Application.DisplayFullScreen = True
    With ActiveWindow
         .DisplayHorizontalScrollBar = False
         .DisplayVerticalScrollBar = False
         .DisplayWorkbookTabs = False
         .DisplayHeadings = False
    End With
    
End Sub
