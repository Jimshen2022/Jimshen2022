Sub save_as()
    '重要代码，如何另存excel文件
    ' save_as Macro
    Range("B1").Select
    ActiveWorkbook.saveas ActiveWorkbook.Path & "\WVG RP TAT MGT-" & Format(Now(), "yyyymmdd.hhmm") & ".xlsx"
    
    Range("C2").Select
    
End Sub