Sub Copy_formula_from_N_to_W() '写公式并将ColumnN~W公式到底,并将item#从N列复制到W列
    
    Application.ScreenUpdating = False
    Dim i&, lrow&
    lrow = Sheet8.Range("b1048576").End(xlUp).row
    With Sheet8
         .Range("n2").Formula = "=ABS(M2)"
         .Range("o2").Formula = "=SUMIFS('TRIP+PO HJ received but KNQ not'!$B:$B,'TRIP+PO HJ received but KNQ not'!$C:$C,DATA!$B:$B,'TRIP+PO HJ received but KNQ not'!$A:$A,DATA!$O$1)"
         .Range("p2").Formula = "=SUMIFS('TRIP+PO HJ received but KNQ not'!$B:$B,'TRIP+PO HJ received but KNQ not'!$C:$C,DATA!$B:$B,'TRIP+PO HJ received but KNQ not'!$A:$A,DATA!$P$1)"
         .Range("q2").Formula = "=SUMIFS('Mapics adjusted'!N:N,'Mapics adjusted'!C:C,DATA!B:B)"
         .Range("r2").Formula = "=SUMIFS(NG!A:A,NG!E:E,DATA!B:B)"
         .Range("s2").Formula = "=SUMIFS('KNQ delared but HJ not'!AB:AB,'KNQ delared but HJ not'!Y:Y,DATA!B:B)"
         .Range("t2").Formula = "=SUMIFS('Trailer in Yard but HJ&KNQ not'!N:N,'Trailer in Yard but HJ&KNQ not'!K:K,DATA!B:B)"
         .Range("u2").Formula = "=M2+O2-P2-S2"
         .Range("v2").Formula = "=COUNTIFS(NG!E:E,DATA!B:B,NG!J:J,""=NG001SC1"")"
         .Range("w2").Formula = "=IF(AND(D2=G2,G2=L2),""MAPICS=HJ=KNQMAN"",IF(AND(D2=G2+H2,G2+H2=L2),""MAPICS=HJ=KNQMAN"",IF(AND(D2=G2+H2,G2+H2<>L2),""MAPICS=HJ<>KNQMAN"",IF(AND(D2=G2,G2<>L2),""MAPICS=HJ<>KNQMAN"",IF(AND(D2<>G2,G2=L2),""MAPICS<>HJ=KNQMAN"",IF(AND(D2=L2,G2<>L2),""MAPICS=KNQMAN<>HJ"",IF(AND(D2<>L2,G2<>L2),""MAPICS<>HJ<>KNQMAN"",""VIEW"")))))))"
        
         .Range("n2:w2").AutoFill Destination:=Range("n2:w" & lrow)
         .Range("x2:x" & lrow).NumberFormat = "@"
         .Range("b2:b" & lrow).Copy Destination:=Range("x2:x" & lrow)
         .Range("x2:x" & lrow).Value = Range("x2:x" & lrow).Value
         .Range("a1048576").End(3).Offset(1, 0).Resize(50000, 39).Clear
        
    End With
    Application.ScreenUpdating = True
    
    
    
End Sub