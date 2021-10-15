VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Login"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4950
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
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   2340
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPW 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox txtUID 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   2085
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   12347
            MinWidth        =   12347
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   3480
      TabIndex        =   0
      Top             =   840
      Width           =   1360
      Begin VB.CommandButton cmdLogin 
         BackColor       =   &H8000000A&
         Caption         =   "&Login"
         Height          =   375
         Left            =   120
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   3
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
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblUserName 
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "System Login"
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
      Left            =   1250
      TabIndex        =   6
      Top             =   480
      Width           =   5295
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Close All
    Unload Me
End Sub
Private Sub cmdLogin_Click()
    UserID = FN_TB_NUMValidation(Len(Trim(txtUID.Text)), "User ID", Val(txtUID.Text))
    If Val(UserID) > 0 Then
        PW = FN_TB_STRValidation(Len(Trim(txtPW.Text)), "Password", Trim(txtPW.Text))
        If Len(PW) > 0 Then
            PR_Login
        Else
            txtPW.SetFocus
        End If
    Else
        txtUID.SetFocus
    End If
End Sub
Public Sub PR_Login()
    STRSQL = "Select * from Inx_sys_Login where U_ID=" & UserID
    PR_SQL_Execution STRSQL
    If RS.EOF = False Then
        If Trim(RS!U_PW) = PW Then
            MDIHome.Show
            Unload Me
            PR_Close_Con
        Else
            MsgBox "Invalid Password", vbCritical, "Login Fail"
            txtPW.Text = ""
        End If
    Else
        MsgBox "UserID NOT found", vbCritical, "Login Fail"
    End If
End Sub
Private Sub txtUID_LostFocus()
    If txtUID.Text <> "" Or txtPW.Text <> "" Then
        UserID = FN_TB_NUMValidation(Len(Trim(txtUID.Text)), "User ID", Val(txtUID.Text))
        STRSQL = "Select * from Inx_sys_Login where U_ID=" & UserID
        PR_SQL_Execution STRSQL
        If RS.EOF = False Then
            lblUserName.Caption = Trim(RS!U_Name)
        Else
            lblUserName.Caption = ""
            MsgBox "User NOT Found", vbExclamation
        End If
        Set RS = New ADODB.Recordset
    End If
End Sub
