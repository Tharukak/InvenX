Attribute VB_Name = "Security"
Public Function MenuAccess(ByRef ModuleId As Integer) As Integer
    'Find User Profile
    Dim URS As ADODB.Recordset
    Set URS = New ADODB.Recordset
    PR_Open_Con
    URS.Open "Select * from Inx_Sec_Profiles Where U_ID=" & UserID & " and Module_ID=" & ModuleId, Con, adOpenStatic, adLockReadOnly
    If URS.EOF = True Then
        MsgBox "User Profile NOT Created, Contact your InvenX Administrator", vbCritical
        RightsMode = 0
    Else
        If URS!Rights = True Then
            RightsMode = 1
        Else
            MsgBox "Access Denied, Contact your InvenX Administrator", vbCritical
            RightsMode = 0
        End If
    End If
    PR_Open_Con
End Function
