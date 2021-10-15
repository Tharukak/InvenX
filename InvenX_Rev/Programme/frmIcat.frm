VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmIcat 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Item Category Master"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7695
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
   Picture         =   "frmIcat.frx":0000
   ScaleHeight     =   9075
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   7815
      Left            =   6240
      TabIndex        =   5
      Top             =   840
      Width           =   1335
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H8000000A&
         Caption         =   "&Delete"
         Height          =   375
         Left            =   120
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   7320
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFCategory 
      Height          =   7335
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   12938
      _Version        =   393216
   End
   Begin VB.TextBox txtCatName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Top             =   960
      Width           =   4215
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8700
      Width           =   7695
      _ExtentX        =   13573
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Catagory Name"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Category Configuration"
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
      Left            =   1245
      TabIndex        =   1
      Top             =   480
      Width           =   5295
   End
End
Attribute VB_Name = "frmIcat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
On Error GoTo er_EH:
    Dim CatName As String
    CatName = UCase(FN_TB_STRValidation(Len(txtCatName.Text), "Item Category", Trim(txtCatName.Text)))
    'Duplicate Checking
    STRSQL = "Select * from Inx_MSTR_Category Where Cat_Name='" & CatName & "'"
    PR_SQL_Execution STRSQL
    
    If RS.RecordCount > 0 Then
        MsgBox "Category Already in the System", vbExclamation
    Else
        If Len(CatName) = 0 Then
           txtCatName.SetFocus
        Else
            STRSQL = "INSERT INTO Inx_MSTR_Category(Cat_Name) VALUES('" & CatName & "')"
            PR_SQL_Execution STRSQL
            PR_FMT_Grid
        End If
    End If
    Exit Sub
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub cmdDelete_Click()
    If MsgBox("Are you sure Do you Want to delete ", vbQuestion + vbYesNo, "Delete Record") = vbYes Then
        If MSFCategory.Rows > -1 Then
            STRSQL = "Delete from Inx_MSTR_Category where Cat_ID=" & Val(MSFCategory.TextMatrix(MSFCategory.Row, 0))
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
    If Trim(txtCatName.Text) = "" Then
        STRSQL = "Select * from Inx_MSTR_Category Order by Cat_ID"
    Else
        STRSQL = "Select * from Inx_MSTR_Category Where Cat_Name like '" & Trim(txtCatName.Text) & "%'"
    End If
    PR_SQL_Execution STRSQL
    With MSFCategory
    .Cols = 2
    .Rows = 1
    R = 1
    Do While RS.EOF = False
        .Rows = .Rows + 1
        .TextMatrix(R, 0) = Trim(RS!Cat_ID)
        .TextMatrix(R, 1) = Trim(RS!Cat_Name)
        R = R + 1
        RS.MoveNext
    Loop
    .TextMatrix(0, 0) = "ID"
    .TextMatrix(0, 1) = "Category Name"
    
    .ColWidth(0) = 800
    .ColWidth(1) = 4000
    End With
End Sub

Private Sub Form_Load()
    PR_FMT_Grid
End Sub
