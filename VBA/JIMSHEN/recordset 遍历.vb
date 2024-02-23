
'recordset 遍历

'     With rstRecordset
'    .MoveLast
'    .MoveFirst
'    For i = 1 To .RecordCount
'        For j = 1 To .Fields.Count
'             tblArray(i, j) = .Fields(j).Value
'        Next j
'        .MoveNext
'    Next i
'End With