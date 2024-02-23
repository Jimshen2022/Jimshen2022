Sub load_UP_from_as400()  'sheets UP
    
    Application.ScreenUpdating = False
    Dim i As Long
    Dim adors As New Recordset
    
    
    Sheet7.Activate
    Sheet7.Cells.Clear
    'Sheet14.Range("a1").CurrentRegion.Copy Sheet7.Range("a1")
    
    Set Db = New Connection
    Db.CursorLocation = adUseClient
    If Db.State = 1 Then Db.Close
    U = Sheet9.Range("a1").Value
    P = Sheet9.Range("a2").Value
    
   
    Db.Open "Provider =IBMDASQL.DataSource.1" & _
     ";Catalog Library List=JIMTDTA" & _
     ";Persist Security Info=True" & _
     ";Force Translate=0" & _
     ";Data Source = 10.9.3.106 " & _
     ";User ID =" & U & "" & _
     ";Password =" & P
          
     Set adors = New Recordset
     If adors.State = 1 Then adors.Close

'    cmdtxt = "SELECT u.ITNBR,u.AVCST/23081 as UP_$USD  " & _
'             "FROM AMFLIBW.ITEMBL u " & _
'             "Where u.HOUSE IN ('35') And u.AVCST<>0 AND u.ITCLS like 'Z%' AND u.ITCLS not like '%K' "
             
        cmdtxt = "SELECT a.RPAITX,(CASE WHEN a.RPBRCD IN ('VND') THEN a.RPAMVA/22685 ELSE a.RPAMVA END) AS RPAMVA,a.RPBLDT,a.RPZ0D7, T2.ITCLS " & _
             "FROM AMFLIBW.ITMFPR a " & _
             "left join AMFLIBW.ITMRVA T2 on a.RPAITX=T2.ITNBR and a.RPZ0D7 = T2.STID " & _
             "WHERE a.RPZ0D7 = '35' AND a.RPAITX||a.RPZ0D7||a.RPBLDT IN " & _
             "(SELECT a.RPAITX||a.RPZ0D7||MAX(a.RPBLDT) RPBLDT " & _
             "FROM AMFLIBW.ITMFPR a  WHERE a.RPZ0D7 = '35' GROUP BY a.RPAITX,a.RPZ0D7) " & _
             "AND T2.ITCLS LIKE 'Z%' "
             

    adors.Open cmdtxt, Db, 3, 3
  
     For i = 0 To adors.Fields.Count - 1
         Sheet7.Cells(1, i + 1) = adors.Fields(i).Name
     Next i
     
     Sheet7.Range("a1048576").End(3).Offset(1, 0).CopyFromRecordset adors
     adors.Close
     Set adors = Nothing
    Sheet7.Columns("A:A").NumberFormat = "@"
    Application.ScreenUpdating = True
    
End Sub