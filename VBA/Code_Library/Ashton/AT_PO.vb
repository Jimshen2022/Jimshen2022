Sub OpenPO666()
    
    
    Dim i As Long
    Dim adors As New Recordset
    Dim ws As Workbook
    
    ws.Activate.clearcontents
    
    Set Db = New Connection
    Db.CursorLocation = adUseClient
    If Db.State = 1 Then Db.Close
    
    
    Db.Open "Provider =IBMDASQL.DataSource.1" &  _
            ";Catalog Library List=JIMTDTA" &  _
            ";Persist Security Info=True" &  _
            ";Force Translate=0" &  _
            ";Data Source = AFIPROD" &  _
            ";User ID =JIMSHEN" &  _
            ";Password =MJ2055"
    
    Set adors = New Recordset
    If adors.State = 1 Then adors.Close
    
    cmdtxt = "SELECT POITEM.ITNBR, POITEM.QTYOR, POMAST.HOUSE, POMAST.ORDNO, POMAST.VNDNR, POMAST.PSTTS,ITMEXT.UUCCIM, POITEM.DUEDT, ITBEXT.ITMCLSID, ITBEXT.PICKPUT, ITBEXT.SCOOPQTY, ITBEXT.SKIDSIZE,VENNAML0.VNNMVM " &  _
            "FROM AFILELIB.ITBEXT ITBEXT, AFILELIB.ITMEXT ITMEXT, AMFLIBA.POITEM POITEM, AMFLIBA.POMAST POMAST, AMFLIBA.VENNAML0 VENNAML0 " &  _
            "WHERE ITBEXT.ITNBR = ITMEXT.ITNBR AND POITEM.ITNBR = ITMEXT.ITNBR AND POITEM.ORDNO = POMAST.ORDNO AND POMAST.VNDNR = VENNAML0.VNDRVM AND POITEM.HOUSE = ITBEXT.HOUSE AND POMAST.HOUSE = POITEM.HOUSE AND ((POMAST.PSTTS='10') OR (POMAST.PSTTS='20') OR (POMAST.PSTTS='30')) AND (POITEM.HOUSE='335') AND (POMAST.VNDNR NOT IN ('600039','900639','900515')) " &  _
            "ORDER BY POITEM.ITNBR, POITEM.DUEDT"
    '
    'SELECT POITEM.ITNBR, POITEM.QTYOR, POMAST.HOUSE, POMAST.ORDNO, POMAST.VNDNR, POMAST.PSTTS,ITMEXT.UUCCIM, POITEM.DUEDT, ITBEXT.ITMCLSID, ITBEXT.PICKPUT, ITBEXT.SCOOPQTY, ITBEXT.SKIDSIZE,VENNAML0.VNNMVM
    'FROM AFILELIB.ITBEXT ITBEXT, AFILELIB.ITMEXT ITMEXT, AMFLIBA.POITEM POITEM, AMFLIBA.POMAST POMAST, AMFLIBA.VENNAML0 VENNAML0
    'WHERE ITBEXT.ITNBR = ITMEXT.ITNBR AND POITEM.ITNBR = ITMEXT.ITNBR AND POITEM.ORDNO = POMAST.ORDNO AND POMAST.VNDNR = VENNAML0.VNDRVM AND POITEM.HOUSE = ITBEXT.HOUSE AND POMAST.HOUSE = POITEM.HOUSE AND ((POMAST.PSTTS='10') OR (POMAST.PSTTS='20') OR (POMAST.PSTTS='30')) AND (POITEM.HOUSE='335') AND (POMAST.VNDNR NOT IN ('600039','900639','900515'))
    'ORDER BY POITEM.ITNBR, POITEM.DUEDT
    
    adors.Open cmdtxt, Db, 3, 3
    For i = 0 To adors.Fields.count - 1
        ws.Sheet1.Cells(1, i + 2) = adors.Fields(i).Name
    Next i
    
    ws.Sheet1.Cells(1, i + 2).Range("B2").CopyFromRecordset adors
    adors.Close
    Set adors = Nothing
    
    
    ws.Sheet1.Activate
    Range("A1") = "Due_Date"
    
    Dim m%
    For m = 2 To [b66365].End(3).Row
        Cells(m, "a") = CDate(Format(Mid(Cells(m, "i"), 2, 6), "0000-00-00"))
        
    Next
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = ture
    Application.StatusBar = False
    
    
End Sub