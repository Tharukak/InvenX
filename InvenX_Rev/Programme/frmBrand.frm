VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBrand 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Brand Master"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7560
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
   Picture         =   "frmBrand.frx":0000
   ScaleHeight     =   9075
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBrand 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1920
      TabIndex        =   5
      Top             =   960
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      Height          =   7815
      Left            =   6120
      TabIndex        =   0
      Top             =   840
      Width           =   1335
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H8000000A&
         Caption         =   "E&xit"
         Height          =   375
         Left            =   120
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   7320
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H8000000A&
         Caption         =   "&Add"
         Height          =   375
         Left            =   120
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H8000000A&
         Caption         =   "&Delete"
         Height          =   375
         Left            =   120
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   720
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFBrand 
      Height          =   7335
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   12938
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   8700
      Width           =   7560
      _ExtentX        =   13335
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Brand Configuration"
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Brand Name"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   1935
   End
End
Attribute VB_Name = "frmBrand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
On Error GoTo er_eh:
    Dim BrandName As String
    BrandName = UCase(FN_TB_STRValidation(Len(txtBrand.Text), "Item Brand", Trim(txtBrand.Text)))
    
    'Duplicate Checking
    STRSQL = "Select * from Inx_MSTR_Brand Where Brand_Name='" & BrandName & "'"
    PR_SQL_Execution STRSQL
    
    If RS.RecordCount > 0 Then
        MsgBox "Brand Already in the System", vbExclamation
    Else
        If Len(BrandName) = 0 Then
            txtBrand.SetFocus
        Else
            STRSQL = "INSERT INTO Inx_MSTR_Brand(Brand_Name) VALUES('" & BrandName & "')"
            PR_SQL_Execution STRSQL
            PR_FMT_Grid
        End If
    End If
    Exit Sub
er_eh:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub cmdDelete_Click()
    If MsgBox("Are you sure Do you Want to delete ", vbQuestion + vbYesNo, "Delete Record") = vbYes Then
        If MSFBrand.Rows > -1 Then
            STRSQL = "Delete from Inx_MSTR_Brand where Brand_ID=" & Val(MSFBrand.TextMatrix(MSFBrand.Row, 0))
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
    If Trim(txtBrand.Text) = "" Then
        STRSQL = "Select * from Inx_MSTR_Brand Order by Brand_ID"
    Else
        STRSQL = "Select * from Inx_MSTR_Brand Where Brand_Name like '" & Trim(txtBrand.Text) & "%'"
    End If
    PR_SQL_Execution STRSQL
    With MSFBrand
        .Cols = 2
        .Rows = 1
        R = 1
        Do While RS.EOF = False
            .Rows = .Rows + 1
            .TextMatrix(R, 0) = Trim(RS!Brand_ID)
            .TextMatrix(R, 1) = Trim(RS!Brand_Name)
            R = R + 1
            RS.MoveNext
        Loop
        .TextMatrix(0, 0) = "ID"
        .TextMatrix(0, 1) = "Brand Name"
        
        .ColWidth(0) = 800
        .ColWidth(1) = 4000
    End With
End Sub

Private Sub Form_Load()
    PR_FMT_Grid
End Sub

