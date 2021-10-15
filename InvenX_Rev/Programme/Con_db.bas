Attribute VB_Name = "Con_db"
Public Sub db_Connect()
    Set Con = New ADODB.Connection
    'BLI --
    Con.ConnectionString = "Provider=SQLOLEDB.1;Password=SinX@123;Persist Security Info=True;User ID=InvenX_User;Initial Catalog=InvenX_Rev;Data Source=10.150.152.17"
    'BIA -- Con.ConnectionString = "Provider=SQLOLEDB.1;Password=BIA@123;Persist Security Info=True;User ID=InvenX_Admin;Initial Catalog=InvenX_Rev;Data Source=BCI-CTSQL-01\BIACTSQL" 'BIA --
End Sub

Public Sub PR_Open_Con()
    db_Connect
    If Con.State = adStateClosed Then Con.Open
End Sub

Public Sub PR_Close_Con()
    db_Connect
    If Con.State = adStateOpen Then Con.Close
End Sub
