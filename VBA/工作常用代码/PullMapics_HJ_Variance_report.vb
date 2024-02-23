Sub Pull_AS_Mapivs_HJ_Variance()
    
    Dim wb As Workbook
    Dim arr, i&
    
    't = Timer
    Application.ScreenUpdating = False
    'Sheet4.Range("a1").Value = "Data collected at:" & Format(Now(), "hhmm,mm-dd-yyyy")
    
    With Sheet3
         .Cells.Clear
        
        Set wb = GetObject("C:\Users\jishen\Downloads\AS_Mapics_HJ.xlsx") '打开工作簿
        'arr = wb.ActiveSheet.Range("a2:k" & Range("a7000").End(3).Row).Value
        arr = wb.ActiveSheet.Range("a2:k10000")
        wb.ActiveSheet.Range("a1").Copy Sheet3.Range("w1")
        wb.Close False
        
        '.Range("a1:k1") = Array("Whse", "Item Number", "Mapics HJ Diff", "Mapics", "Not RM'd", "Shipped Not Invoiced", "WA", "YA", "In Yard No RTS", "Transfer", "Loaded Not HJ")
         .Columns("a:k").NumberFormat = "@"
         .Range("a1").Resize(UBound(arr), UBound(arr, 2)).Value = arr
         .Range("c2:k10000").NumberFormatLocal = ""
         .Range("c2:k10000").Value = Range("c2:k10000").Value
         .Columns("c:c").Insert Shift:=xlToRight
         .Range("c1").Value = "Reason"
         .Columns("a:k").EntireColumn.AutoFit
         .Columns("c:c").ColumnWidth = 40
    End With
    Application.ScreenUpdating = True
    'MsgBox Format(Timer - t, "0.00" & "s")
    
End Sub
