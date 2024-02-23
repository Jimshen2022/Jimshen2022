Sub generate_sht_list_()
	Dim sht As Worksheet
	Dim n%
	n = 1
	Cells.Clear
    For Each sht In Worksheets
        If sht.Name <> ActiveSheet.Name Then
            Cells(n, 1) = sht.Name
            ActiveSheet.Hyperlinks.Add anchor:=Cells(n, 2), Address:="", SubAddress:="'" & sht.Name & "'!a1", TextToDisplay:=sht.Name & "link"
            n = n + 1
        End If
    Next
End Sub