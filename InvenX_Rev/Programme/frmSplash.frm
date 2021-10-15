VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "InvenX Revolution"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11730
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
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   5805
   ScaleWidth      =   11730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   2520
      Top             =   3360
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   375
      Left            =   -240
      TabIndex        =   0
      Top             =   5400
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   2880
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "InvenX"
      BeginProperty Font 
         Name            =   "Neuropolitical Rg"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1095
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   4935
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    lblVersion.Caption = "Revolution Version - " & App.Major & "." & App.Minor & "," & App.Revision
End Sub

Private Sub Timer1_Timer()
    If PB.Value < 100 Then
        PB.Value = PB.Value + 1
    Else
        Load frmLogin
        frmLogin.Show
        Unload Me
    End If
End Sub
