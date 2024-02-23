Sub LoadItemAndUP()

't = Timer
'Application.ScreenUpdating = False
'Application.Calculation = xlCalculationManual
'Application.StatusBar = "Loading UOM, please wait ......"

Sheet14.Cells.ClearContents

Set Db = New Connection
    Db.CursorLocation = adUseClient
    If Db.State = 1 Then Db.Close

UserID = Sheet25.Range("a1").Value
PW = Sheet25.Range("a2").Value

    Db.Open "Provider =IBMDASQL.DataSource.1" & _
     ";Catalog Library List=JDETSTDTA" & _
     ";Persist Security Info=True" & _
     ";Force Translate=0" & _
     ";Data Source = AFIPROD" & _
     ";User ID = " & UserID & "" & _
     ";Password = " & PW

     Set adors = New Recordset
     If adors.State = 1 Then adors.Close

    cmdtxt = "SELECT ITMRVA.ITNBR, ITEMBL.HOUSE, ITEMBL.ITCLS, ITMRVA.B2Z95S, ITEMBL.MOHTQ, ITMRVA.ITDSC, ITBEXT.ITMCLSID, ITBEXT.PICKPUT, ITBEXT.SCOOPQTY, ITBEXT.SKIDSIZE " & _
             "FROM AFILELIB.ITBEXT ITBEXT, D20ACF9V.AMFLIBA.ITEMBL ITEMBL, D20ACF9V.AFILELIB.ITMEXT ITMEXT, D20ACF9V.AMFLIBA.ITMRVA ITMRVA, D20ACF9V.AMFLIBA.WHSMST WHSMST " & _
             "WHERE ITMRVA.ITNBR = ITEMBL.ITNBR AND ITMRVA.STID = WHSMST.STID AND ITBEXT.HOUSE = ITEMBL.HOUSE AND ITBEXT.ITNBR = ITEMBL.ITNBR AND ITBEXT.ITNBR = ITMRVA.ITNBR AND ITMEXT.ITNBR = ITBEXT.ITNBR AND ITMEXT.ITNBR = ITEMBL.ITNBR AND ITMEXT.ITNBR = ITMRVA.ITNBR AND ITEMBL.HOUSE = WHSMST.WHID AND ((ITEMBL.HOUSE='335')) " & _
             "ORDER BY ITMRVA.ITNBR,ITBEXT.SCOOPQTY"

    adors.Open cmdtxt, Db, 3, 3
     For i = 0 To adors.Fields.Count - 1
         Sheet14.Cells(1, i + 1) = adors.Fields(i).Name
     Next i

     Sheet14.Range("a2").CopyFromRecordset adors
     adors.Close
     Set adors = Nothing
    
    
    
     Sheet14.Activate
     Columns("a:j").NumberFormat = "@"
     Range("k1") = "CBM"
    
     Dim m&, nrow&, arr()
     nrow = Range("a1048576").End(3).Row
     arr = Range("d2:k" & nrow)


    For m = 1 To nrow - 1
        arr(m, 8) = arr(m, 1) * 0.028317
    Next m
    Range("d2:k" & m) = arr
    

    
    Sheet24.Activate
    Dim x&, xrow&
    xrow = Sheet24.Range("a1048576").End(3).Row
    Range("a2:k" & xrow).Copy Sheet14.Range("a1048576").End(3).Offset(1, 0)
    
    
    Call LoadATUP
    'ActiveWorkbook.Save
    Sheet15.Select


'    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True
'    Application.StatusBar = False
'MsgBox Format(Timer - t, "0.00") & "s"


End Sub


Sub LoadATUP()


Sheet15.Activate
Sheet15.Cells.ClearContents

Set Db = New Connection
    Db.CursorLocation = adUseClient
    If Db.State = 1 Then Db.Close
    
    U = Sheet25.Range("a1")
    P = Sheet25.Range("a2")
    Db.Open "Provider =IBMDASQL.DataSource.1" & _
     ";Catalog Library List=JDETSTDTA" & _
     ";Persist Security Info=True" & _
     ";Force Translate=0" & _
     ";Data Source = AFIPROD" & _
     ";User ID =" & U & "" & _
     ";Password = " & P

     Set adors = New Recordset
     If adors.State = 1 Then adors.Close

'        cmdtxt = "SELECT u.ITNO1G,u.UCST1G as UP_$USD  " & _
'                 "FROM AMFLIBA.ITMPRB u " & _
'                 "Where u.STID1G IN ('335') and u.UCST1G<>0 "


'        cmdtxt = "SELECT u.ITNO1G,u.UCST1G as UP_$USD,t1.itcls  " & _
'                 "FROM AMFLIBA.ITMPRB u,AMFLIBA.ITMRVA t1 " & _
'                 "Where u.STID1G IN ('335') and u.UCST1G<>0 and u.ITNO1G = t1.itnbr and u.STID1G = t1.stid and t1.itcls like 'Z%' AND t1.itcls not like '%K' "
'

    cmdtxt = "SELECT PRICE.PITEM, PRICE.PAMNT " & _
             "FROM AFILELIB.PRICE PRICE " & _
             "WHERE PRICE.PRICCD='FOBARC' " & _
             "ORDER BY PRICE.PITEM"

    adors.Open cmdtxt, Db, 3, 3
     For i = 0 To adors.Fields.Count - 1
         Sheet15.Cells(1, i + 1) = adors.Fields(i).Name
     Next i
     
     Sheet15.Range("a2").CopyFromRecordset adors
     adors.Close
     Set adors = Nothing


End Sub