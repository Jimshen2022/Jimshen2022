Sub interior_()
	On Error Resume Next
	Application.DisplayAlerts = False
	Dim i%, j%

		For j = 1 To Worksheets.Count
			If Worksheets(j).Name = "ColorIndex" Then
				Worksheets(j).Cells.Clear
			End If
		Next

		With Worksheets("ColorIndex")
			For i = 1 To 100
				.Range("a" & i).Interior.ColorIndex = i
				.Range("b" & i).Value = i
				
			Next
		End With
	Application.DisplayAlerts = True
End Sub