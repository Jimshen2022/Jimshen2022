Sub Copy_AT_CC_WKQS_ColumnA_to_F() 'copy ColumnA~F公式到G列
    Dim i&, nrow&, m&, lrow&
    Sheet4.Select
    lrow = Sheet4.Range("g1048576").End(xlUp).Row
    Range("a2:f2").AutoFill Destination:=Range("a2:f" & lrow)
    
End Sub


Sub SKUcount()    
    
    Dim i%, nrow%
    nrow = Sheet4.Range("m1047586").End(3).Row
    
    With Sheet4
         .Range("h2:i" & nrow).NumberFormat = "General"
         .Range("i2") = "=IF(COUNTIFS($M$2:M2,M2)>1,0,COUNTIFS($M$2:M2,M2))"
         .Range("h2") = "=IF(OR(LEFT(M2,1)=""A"",LEFT(M2,1)=""L"",LEFT(M2,1)=""Q"",LEFT(M2,1)=""R""),""Accessory"",IF(OR(LEFT(M2,4)=""100-"",LEFT(M2,1)=""B"",LEFT(M2,1)=""D"",LEFT(M2,1)=""H"",LEFT(M2,1)=""T"",LEFT(M2,1)=""W"",LEFT(M2,1)=""P"",LEFT(M2,1)=""Z""),""CG"",""UPH"")) "
         .Range("h2:i2").AutoFill Destination:=Range("h2:i" & nrow)
        
    End With
    
End Sub