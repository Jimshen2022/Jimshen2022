Sub GetTypeList()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    PT1 = "V:\Prod & Inv Control\Public\00.Master Data\UPHMaster.accdb"
    'On Error Resume Next
line1:
    'Get TYPE LIST
    
    Set x = New Connection
    x.CursorLocation = adUseClient
    If x.State = 1 Then x.Close
    
    x.Open "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" & PT1
    
    Set Y = New Recordset
    If Y.State = 1 Then Y.Close
    
    SQL = "select * from [Type] "
    
    Y.Open SQL, x, 3, 3
    
    For J = 0 To Y.Fields.Count - 1
        Worksheets("UPHData").Cells(1, J + 1) = Y.Fields(J).Name
    Next J
    
    Worksheets("UPHData").Range("A2").CopyFromRecordset Y
    Y.Close
    x.Close
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub