Sub Cubes() 'ready 5.1
    
    Dim i As Long
    Dim adors As New Recordset
    
    Worksheets("Cubes").Range("A2:D66365").ClearContents
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Loading Unit Cubes, please wait ......"
    
    Set Db = New Connection
    Db.CursorLocation = adUseClient
    If Db.State = 1 Then Db.Close
    
    
    Db.Open "Provider =IBMDASQL.DataSource.1" &  _
            ";Catalog Library List=JIMTDTA" &  _
            ";Persist Security Info=True" &  _
            ";Force Translate=0" &  _
            ";Data Source = AFIPROD" &  _
            ";User ID =JIMSHEN" &  _
            ";Password =MJ2056"
    
    Set adors = New Recordset
    If adors.State = 1 Then adors.Close
    
    cmdtxt = "SELECT ITMRVA.ITNBR, ITEMBL.HOUSE, ITEMBL.ITCLS, ITMRVA.B2Z95S as UnitCube " &  _
            "FROM D20ACF9V.AFILELIB.ITBEXT ITBEXT, D20ACF9V.AMFLIBA.ITEMBL ITEMBL, D20ACF9V.AFILELIB.ITMEXT ITMEXT, D20ACF9V.AMFLIBA.ITMRVA ITMRVA, D20ACF9V.AMFLIBA.WHSMST WHSMST " &  _
            "WHERE ITMRVA.ITNBR = ITEMBL.ITNBR AND ITMRVA.STID = WHSMST.STID AND ITBEXT.HOUSE = ITEMBL.HOUSE AND ITBEXT.ITNBR = ITEMBL.ITNBR AND ITBEXT.ITNBR = ITMRVA.ITNBR AND ITMEXT.ITNBR = ITBEXT.ITNBR AND ITMEXT.ITNBR = ITEMBL.ITNBR AND ITMEXT.ITNBR = ITMRVA.ITNBR AND ITEMBL.HOUSE = WHSMST.WHID AND (ITEMBL.HOUSE='335') and ITEMBL.ITCLS LIKE 'Z%' AND ITEMBL.ITCLS NOT LIKE '%K' " &  _
            "ORDER BY ITMRVA.ITNBR"
    
    'SELECT ITMRVA.ITNBR, ITEMBL.HOUSE, ITEMBL.ITCLS, ITMRVA.B2Z95S as UnitCube
    'FROM D20ACF9V.AFILELIB.ITBEXT ITBEXT, D20ACF9V.AMFLIBA.ITEMBL ITEMBL, D20ACF9V.AFILELIB.ITMEXT ITMEXT, D20ACF9V.AMFLIBA.ITMRVA ITMRVA, D20ACF9V.AMFLIBA.WHSMST WHSMST
    'WHERE ITMRVA.ITNBR = ITEMBL.ITNBR AND ITMRVA.STID = WHSMST.STID AND ITBEXT.HOUSE = ITEMBL.HOUSE AND ITBEXT.ITNBR = ITEMBL.ITNBR AND ITBEXT.ITNBR = ITMRVA.ITNBR AND ITMEXT.ITNBR = ITBEXT.ITNBR AND ITMEXT.ITNBR = ITEMBL.ITNBR AND ITMEXT.ITNBR = ITMRVA.ITNBR AND ITEMBL.HOUSE = WHSMST.WHID AND (ITEMBL.HOUSE='335') and ITEMBL.ITCLS LIKE 'Z%' AND ITEMBL.ITCLS NOT LIKE '%K'
    'ORDER BY ITMRVA.ITNBR
    
    adors.Open cmdtxt, Db, 3, 3
    For i = 0 To adors.Fields.Count - 1
        Worksheets("Cubes").Cells(1, i + 1) = adors.Fields(i).Name
    Next i
    
    Worksheets("Cubes").Range("A2").CopyFromRecordset adors
    adors.Close
    Set adors = Nothing
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    
    
End Sub