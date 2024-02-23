
Sub 文本改数值()
    
    Sheet10.Range("zz1") = 1
    Range("zz").Copy
    Range("B:B").Select #这里需要限定下界 ，否则遇到空白单元格会变成0
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlMultiply, SkipBlanks:=False, Transpose:=False
    range("zz1").ClearContents
    
End Sub


Sub gs2()
    
    '
    Range("G1") = 1
    Range("G1").Select
    Selection.Copy
    Range("B2:B100").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlMultiply, SkipBlanks:=False, Transpose:=False
    
    
    Range("b2:b100").NumberFormat = "##,##0.00"
    .Range("D2:D12").NumberFormat = "##,##0"   '千位符
	
	
    '   Format用法()
    '   MsgBox Format("123", "000000")
    '   MsgBox Format(Date, "M/d/yy")
    '   MsgBox Format(Date, "yyyy-mm-dd")
    '   MsgBox Format(Date, "yyyy年mm月yy日")
    '   MsgBox Format(Time, "hh:mm:ss")
    '   MsgBox Format("0.5556", "0.00")
    '   MsgBox Format("aaa", ">")
    '   MsgBox Format("AAA", "<")
    
End Sub
