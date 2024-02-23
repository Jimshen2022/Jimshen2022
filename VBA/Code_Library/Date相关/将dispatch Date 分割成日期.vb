Sub End399()
    Rem 将dispatch Date 分割成日期
    Application.ScreenUpdating = False
    Sheets("DATA").Activate
    For i = 2 To [b1048576].End(3).Row
        '1 [A65536].End(xlUp).Row 'A列末行向上第一个有值的行数 ,xlUp可用3代替,   即向上
        '2 [A1].End(xlDown).Row 'A列首行向下第一个有值之行数 ,xlDown可用4代替,   即向下
        '3 [IV1].End(xlToLeft).Column '第一行末列向左第一列有数值之列数,xlToLeft可用1代替,  即向左
        '3 [A1].End(xlToRight).Column '第一行首列向右有连续值的末列之列数 ,xlToRight可用2代替,   即向右
        
        If Cells(i, "B") <> "" Then Cells(i, "U") = Split(Cells(i, "B"), " ")(0)
    Next i
End Sub