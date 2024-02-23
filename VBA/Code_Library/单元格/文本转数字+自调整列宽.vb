


    With Sheet13.Range("b2:b" & crow, "ak2:ak" & crow)
        .NumberFormatLocal = "0"
        .Value = .Value

        .Columns("a:ak").EntireColumn.AutoFit
        .Range("a1").EntireColumn.AutoFit
        .Range("b1").EntireRow.AutoFill
    End With