-- VBA Product separated:
Sub MILProdctseparated
      For j = 2 To UBound(brr)
            If brr(j, 3) = "TAF" Then
                brr(j, 9) = "RP"
            ElseIf brr(j, 3) = "PACS" Then
                brr(j, 9) = "UnKits"
            ElseIf brr(j, 3) Like "Z*" And brr(j, 3) Like "*K" Then
                brr(j, 9) = "UnKits"
            ElseIf brr(j, 3) = "ZACM" Or brr(j, 3) = "ZASU" Or brr(j, 3) = "ZMLH" Or brr(j, 3) = "ZMLR" Or brr(j, 3) = "ZUSR" Or brr(j, 3) = "ZUSU" Or brr(j, 3) = "ZVUC" Or brr(j, 3) = "ZXUC" Or brr(j, 3) = "ZUSU" Or brr(j, 3) = "ZUMU" Then
                brr(j, 9) = "UPH"
            ElseIf brr(j, 3) = "ZDAA" Or brr(j, 3) = "ZDAY" Or brr(j, 3) = "ZVAA" Or brr(j, 3) = "ZDAB" Or brr(j, 3) = "ZDAW" Or brr(j, 3) = "ZDYB" Then
                brr(j, 9) = "CG"
            ElseIf brr(j, 3) = "ZKIS" Then
                brr(j, 9) = "Bedding"
            ElseIf brr(j, 3) = "WPLS" Then
                brr(j, 9) = "Plastics"
            ElseIf brr(j, 3) = "WVBC" Or brr(j, 3) = "WVCS" Then
                brr(j, 9) = "Foundation"
            ElseIf brr(j, 3) = "PANL" Then
                brr(j, 9) = "Panel"
            ElseIf brr(j, 3) = "ZKIZ" Then
                brr(j, 9) = "ZipperCover"
            ElseIf brr(j, 3) = "BBFR" Or brr(j, 3) = "WVHC" Then
                brr(j, 9) = "Verona"
            ElseIf Not brr(j, 3) Like "Z*" Then
                brr(j, 9) = "RawMaterial"
            Else
                brr(j, 9) = "Check"
            End If
        Next
End Sub