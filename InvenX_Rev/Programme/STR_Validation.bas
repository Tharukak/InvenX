Attribute VB_Name = "STR_Validation"
Public Function FN_TB_NUMValidation(Tb_Length, Sub_Name, Txt_Value) As Integer
    If Tb_Length = 0 Then
        MsgBox Sub_Name & " NOT Entered", vbExclamation
    Else
       FN_TB_NUMValidation = Txt_Value
    End If
End Function

Public Function FN_TB_STRValidation(Tb_Length, Sub_Name, Txt_Text) As String
    If Tb_Length = 0 Then
        MsgBox Sub_Name & " NOT Entered", vbExclamation
    Else
       FN_TB_STRValidation = Txt_Text
    End If
End Function

