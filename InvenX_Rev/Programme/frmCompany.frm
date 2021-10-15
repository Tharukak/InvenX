VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCompany 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Company Master"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7665
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCompany.frx":0000
   ScaleHeight     =   9255
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   6375
      Left            =   6240
      TabIndex        =   15
      Top             =   2400
      Width           =   1335
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H8000000A&
         Caption         =   "&Delete"
         Height          =   375
         Left            =   120
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H8000000A&
         Caption         =   "&Add"
         Height          =   375
         Left            =   120
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H8000000A&
         Caption         =   "E&xit"
         Height          =   375
         Left            =   120
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   5880
         Width           =   1095
      End
   End
   Begin VB.TextBox txtComCode 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6360
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtCont2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4560
      TabIndex        =   5
      Top             =   2040
      Width           =   2895
   End
   Begin VB.TextBox txtCont1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      Top             =   2040
      Width           =   2535
   End
   Begin VB.TextBox txtAdd2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Top             =   1680
      Width           =   6015
   End
   Begin VB.TextBox txtAdd1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   1320
      Width           =   6015
   End
   Begin VB.TextBox txtComName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   960
      Width           =   3615
   End
   Begin MSFlexGridLib.MSFlexGrid MSFCompany 
      Height          =   6375
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   11245
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   8880
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14111
            MinWidth        =   14111
         EndProperty
      EndProperty
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "CompanyCode"
      Height          =   255
      Left            =   5280
      TabIndex        =   14
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "2."
      Height          =   255
      Left            =   4200
      TabIndex        =   13
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No. 1"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Address Line 2"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Address Line 1"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Configuration"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   480
      Width           =   5295
   End
End
Attribute VB_Name = "frmCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
On Error GoTo er_eh:
    Dim ComCode, ComName, ComAdd1, ComAdd2, ComCont1, ComCont2 As String
    ComCode = UCase(FN_TB_STRValidation(Len(txtComCode.Text), "Company Code", Trim(txtComCode.Text)))
    If Len(ComCode) = 0 Then
        txtComCode.SetFocus
        Exit Sub
    End If
    ComName = UCase(FN_TB_STRValidation(Len(txtComName.Text), "Company Name", Trim(txtComName.Text)))
    If Len(ComName) = 0 Then
        txtComName.SetFocus
        Exit Sub
    End If
    ComAdd1 = UCase(FN_TB_STRValidation(Len(txtAdd1.Text), "Company Address Line 01", Trim(txtAdd1.Text)))
    If Len(ComAdd1) = 0 Then
        txtAdd1.SetFocus
        Exit Sub
    End If
    ComAdd2 = UCase(FN_TB_STRValidation(Len(txtAdd2.Text), "Company Address Line 02", Trim(txtAdd2.Text)))
    If Len(ComAdd2) = 0 Then
        txtAdd2.SetFocus
        Exit Sub
    End If
    ComCont1 = UCase(FN_TB_STRValidation(Len(txtCont1.Text), "Company Contact Number 01", Trim(txtCont1.Text)))
    If Len(ComCont1) = 0 Then
        txtCont1.SetFocus
        Exit Sub
    End If
    ComCont2 = UCase(FN_TB_STRValidation(Len(txtCont2.Text), "Company Contact Number 02", Trim(txtCont2.Text)))
    If Len(ComCont2) = 0 Then
        txtCont2.SetFocus
        Exit Sub
    End If
    
    'Duplicate Checking
    STRSQL = "Select * from Inx_MSTR_Company Where Com_Code='" & ComCode & "'"
    PR_SQL_Execution STRSQL
    If RS.RecordCount > 0 Then
        MsgBox "Company Code Already Exist in the System", vbExclamation
        txtComCode.SetFocus
        Exit Sub
    End If
    
    STRSQL = "Select * from Inx_MSTR_Company Where Com_Code='" & ComCode & "' and Com_Name='" & Trim(txtComName.Text) & "'"
    PR_SQL_Execution STRSQL
    If RS.RecordCount > 0 Then
        MsgBox "Company Name Already Exist in the System Under Entered Company Code", vbExclamation
        txtComName.SetFocus
        Exit Sub
    End If
    

    STRSQL = "INSERT INTO Inx_MSTR_Company(Com_Code,Com_Name,Com_Add1,Com_Add2,Com_Cont1,Com_Cont2) VALUES('" & ComCode & "','" & ComName & "','" & ComAdd1 & "','" & ComAdd2 & "','" & ComCont1 & "','" & ComCont2 & "')"
    PR_SQL_Execution STRSQL
    PR_FMT_Grid


    Exit Sub
er_eh:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub cmdDelete_Click()
    If MsgBox("Are you sure Do you Want to delete ", vbQuestion + vbYesNo, "Delete Record") = vbYes Then
        If MSFCompany.Rows > -1 Then
            STRSQL = "Delete from Inx_MSTR_company where Com_Code='" & Trim(MSFCompany.TextMatrix(MSFCompany.Row, 0)) & "'"
            PR_SQL_Execution STRSQL
            PR_FMT_Grid
        Else
            MsgBox "There is No Records to Delete", vbExclamation
        End If
    End If
End Sub

Private Sub cmdExit_Click()
    Close All
    Unload Me
End Sub

Public Sub PR_FMT_Grid()
    If Trim(txtComName.Text) = "" Then
        STRSQL = "Select * from Inx_MSTR_Company Order by Com_Code"
    Else
        STRSQL = "Select * from Inx_MSTR_Company Where Com_Name like '" & Trim(txtComName.Text) & "%'"
    End If
    PR_SQL_Execution STRSQL
    With MSFCompany
        .Cols = 6
        .Rows = 1
        R = 1
        Do While RS.EOF = False
            .Rows = .Rows + 1
            .TextMatrix(R, 0) = Trim(RS!Com_Code)
            .TextMatrix(R, 1) = Trim(RS!Com_Name)
            .TextMatrix(R, 2) = Trim(RS!Com_Add1)
            .TextMatrix(R, 3) = Trim(RS!Com_Add2)
            .TextMatrix(R, 4) = Trim(RS!Com_Cont1)
            .TextMatrix(R, 5) = Trim(RS!Com_Cont2)
            R = R + 1
            RS.MoveNext
        Loop
        .TextMatrix(0, 0) = "Code"
        .TextMatrix(0, 1) = "Company Name"
        .TextMatrix(0, 2) = "Address Line 1"
        .TextMatrix(0, 3) = "Address Line 2"
        .TextMatrix(0, 4) = "Contact No.1"
        .TextMatrix(0, 5) = "Contact No.2"
        
        .ColWidth(0) = 800
        .ColWidth(1) = 2000
        .ColWidth(2) = 2000
        .ColWidth(3) = 2000
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
    End With
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    PR_FMT_Grid
End Sub


