'实例1  普通常见的求不重复值问题

Sub find_unique_items()

t1 = Timer
Application.ScreenUpdating = False

    Dim i&, Myr&, Arr, wb As Workbook
    Dim d, k, t
    Set d = CreateObject("Scripting.Dictionary")
    
    Set wb = CreateObject("C:\Users\jishen\Downloads\WANEK3368W2.xlsx")
    Myr = wb.ActiveSheet.[a1048576].End(xlUp).Row
    Arr = wb.ActiveSheet.Range("a2:a" & Myr)
    wb.Close False
    
    Debug.Print UBound(Arr)
    
    For i = 2 To UBound(Arr)
        d(Arr(i, 1)) = d(Arr(i, 1)) + 1     '将A列相同的累加
    Next
    k = d.keys
    t = d.items
    
    With Sheet2
        .Cells.Clear
        .Range("a2").Resize(d.Count, 1) = Application.Transpose(k)
        .Range("b2").Resize(d.Count, 1) = Application.Transpose(t)
        .Range("a1").Resize(1, 2) = Array("姓名", "重复个数")
            
   End With
   Set d = Nothing
   Erase Arr, k, t

Application.ScreenUpdating = True
MsgBox "it took " & Format(Timer - t1, "###.00") & "s"
ThisWorkbook.Save

End Sub

