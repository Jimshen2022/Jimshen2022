Sub BOM()
    'On Error Resume Next
    Application.ScreenUpdating = False
    Dim i As Integer, j As Integer, n As Integer, m As Integer
    Dim cmdtxt As String
    Dim adors As New Recordset
    Dim sh As Worksheet
    Dim adoCN As Object
    Dim strSQL As String
    Dim objPivotCache As Object
    
    
    Sheets("FG_BOM").Select
    Columns("A:K").Select
    Selection.ClearContents
    Call MakeString
    
    Set Db = New Connection
    Db.CursorLocation = adUseClient
    
    If Db.State = 1 Then Db.Close
    UName = Range("B5")
    UPass = Range("B6")
    
    Db.Open "Provider =IBMDASQL.DataSource.1" &  _
            ";Catalog Library List=JDETSTDTA" &  _
            ";Persist Security Info=True" &  _
            ";Force Translate=0" &  _
            ";Data Source = MILPROD" &  _
            ";User ID =" & UName & "" &  _
            ";Password =" & UPass
    
    
    
    Worksheets("SETTING").Select
    
    sd = Worksheets("SETTING").Range("B2").Text
    
    
    Worksheets("FG_BOM").Select
    Set adors = New Recordset
    If adors.State = 1 Then adors.Close
    
    cmdtxt = "SELECT BOMPIT,BOMCIT,BXDCOMPONENTITEMDESCRIPTION as C_DES,BOMGQT,BOMNQT,BOMRAT,BOMSEQ,BXDUNITOFMEASURE as UNIT,BXDITEMTYPECODE as TYPE,BOMPCL,BOMCCL " &  _
            "FROM RGNFILL.PSTBOMD  " &  _
            "WHERE BOMPIT IN " & sd & " AND BOMGQT <>0 AND BXDITEMTYPECODE <> 0 " &  _
            "GROUP BY BOMPIT,BOMCIT,BXDCOMPONENTITEMDESCRIPTION,BOMGQT,BOMNQT,BOMRAT,BOMSEQ,BXDUNITOFMEASURE,BXDITEMTYPECODE,BOMPCL,BOMCCL " &  _
            "ORDER BY BOMPIT"
    
    adors.Open cmdtxt, Db, 3, 3
    
    For i = 0 To adors.Fields.Count - 1
        Worksheets("FG_BOM").Cells(1, i + 1) = adors.Fields(i).Name
    Next i
    
    
    Worksheets("FG_BOM").Range("A2").CopyFromRecordset adors
    
    Columns("A:K").AutoFit
    Columns("A:K").Select
    With Selection.Font
         .Name = "Arial"
         .Size = 10
         .Strikethrough = False
         .Superscript = False
         .Subscript = False
         .OutlineFont = False
         .Shadow = False
         .Underline = xlUnderlineStyleNone
         .ThemeColor = xlThemeColorLight1
         .TintAndShade = 0
         .ThemeFont = xlThemeFontNone
    End With
    Range("A1").Select
    
    Worksheets("FG_BOM").Select
    Application.ScreenUpdating = True
    MsgBox "FG_BOM Downloaded!"
    
    
End Sub