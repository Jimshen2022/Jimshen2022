

Sub Refresh_All_Data_Connections()
    
    '重要代码，在refresh all以后，再执行其它程序的功能
    For Each cnct In ThisWorkbook.Connections
        Select Case cnct. Type
            Case xlConnectionTypeODBC
                cnct.ODBCConnection.BackgroundQuery = False
            Case xlConnectionTypeOLEDB
                cnct.OLEDBConnection.BackgroundQuery = False
        End Select
    Next cnct
End Sub





























