Sub SKUcount() '单条件汇总,将transfer sheet的数量查到data sheet的Transfer Demand---K栏
    
    Dim i%, nrow%
    nrow = Sheet4.Range("m1047586").End(3).Row
    
    With Sheet4
         .Range("h2:i" & nrow).NumberFormat = "General" '将格式改为一般，否则下公式会表现为文本而无法起作用
         .Range("i2") = "=IF(COUNTIFS($M$2:M2,M2)>1,0,COUNTIFS($M$2:M2,M2))"
         .Range("h2") = "=IF(OR(LEFT(M2,1)=""A"",LEFT(M2,1)=""L"",LEFT(M2,1)=""Q"",LEFT(M2,1)=""R""),""Accessory"",IF(OR(LEFT(M2,4)=""100-"",LEFT(M2,1)=""B"",LEFT(M2,1)=""D"",LEFT(M2,1)=""H"",LEFT(M2,1)=""T"",LEFT(M2,1)=""W"",LEFT(M2,1)=""P"",LEFT(M2,1)=""Z""),""CG"",""UPH"")) "
         .Range("h2:i2").AutoFill Destination:=Range("h2:i" & nrow) '复制公式管齐，这里必须要包括来源的范围
        
    End With
    
End Sub

