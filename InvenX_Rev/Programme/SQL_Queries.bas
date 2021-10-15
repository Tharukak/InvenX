Attribute VB_Name = "SQL_Queries"
Public Sub PR_SQL_Execution(STRSQL)
    PR_Open_Con
    Set RS = New ADODB.Recordset
    RS.Open STRSQL, Con, adOpenStatic, adLockReadOnly
    PR_Close_Con
End Sub

