VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSBU 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Factory Master"
   ClientHeight    =   9615
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7680
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
   Picture         =   "frmSBU.frx":0000
   ScaleHeight     =   9615
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSDataListLib.DataCombo dcmbComName 
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   840
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin VB.TextBox txtSBUName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Top             =   1200
      Width           =   3615
   End
   Begin VB.TextBox txtAdd1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      TabIndex        =   5
      Top             =   1560
      Width           =   6015
   End
   Begin VB.TextBox txtAdd2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      TabIndex        =   6
      Top             =   1920
      Width           =   6015
   End
   Begin VB.TextBox txtCont1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      TabIndex        =   7
      Top             =   2280
      Width           =   2535
   End
   Begin VB.TextBox txtCont2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4560
      TabIndex        =   8
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox txtSBUCode 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6480
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   6615
      Left            =   6240
      TabIndex        =   1
      Top             =   2640
      Width           =   1335
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H8000000A&
         Caption         =   "E&xit"
         Height          =   375
         Left            =   120
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   6120
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H8000000A&
         Caption         =   "&Add"
         Height          =   375
         Left            =   120
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   9
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
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFSBU 
      Height          =   6495
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   11456
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   9240
      Width           =   7680
      _ExtentX        =   13547
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
   Begin VB.Label lblComCode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6480
      TabIndex        =   22
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Code"
      Height          =   255
      Left            =   5280
      TabIndex        =   21
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SBU Configuration"
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
      TabIndex        =   19
      Top             =   480
      Width           =   5295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SBU Name"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Address Line 1"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Address Line 2"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No. 1"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "2."
      Height          =   255
      Left            =   4200
      TabIndex        =   14
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "SBU Code"
      Height          =   255
      Left            =   5280
      TabIndex        =   13
      Top             =   1200
      Width           =   1095
   End
End
Attribute VB_Name = "frmSBU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
On Error GoTo er_eh:
    Dim SBUCode, SBUName, ComCode, ComName, SBUAdd1, SBUAdd2, SBUCont1, SBUCont2 As String
    ComName = UCase(FN_TB_STRValidation(Len(dcmbComName.Text), "Company Name", Trim(dcmbComName.Text)))
    If Len(ComName) = 0 Then
        dcmbComName.SetFocus
        Exit Sub
    End If
    
    STRSQL = "Select * from Inx_MSTR_Company Where Com_Name='" & ComName & "'"
    PR_SQL_Execution STRSQL
    If RS.EOF = False Then
        ComCode = Trim(RS!Com_Code)
    Else
        MsgBox "Company NOT Found", vbExclamation
        Exit Sub
    End If
    
    SBUName = UCase(FN_TB_STRValidation(Len(txtSBUName.Text), "SBU Name", Trim(txtSBUName.Text)))
    If Len(SBUName) = 0 Then
        txtSBUName.SetFocus
        Exit Sub
    End If
    
    SBUCode = UCase(FN_TB_STRValidation(Len(txtSBUCode.Text), "SBU Code", Trim(txtSBUCode.Text)))
    If Len(SBUCode) = 0 Then
        txtSBUCode.SetFocus
        Exit Sub
    End If
    
    SBUAdd1 = UCase(FN_TB_STRValidation(Len(txtAdd1.Text), "SBU Address Line 01", Trim(txtAdd1.Text)))
    If Len(SBUAdd1) = 0 Then
        txtAdd1.SetFocus
        Exit Sub
    End If
    SBUAdd2 = UCase(FN_TB_STRValidation(Len(txtAdd2.Text), "SBU Address Line 02", Trim(txtAdd2.Text)))
    If Len(SBUAdd2) = 0 Then
        txtAdd2.SetFocus
        Exit Sub
    End If
    SBUCont1 = UCase(FN_TB_STRValidation(Len(txtCont1.Text), "SBU Contact Number 01", Trim(txtCont1.Text)))
    If Len(SBUCont1) = 0 Then
        txtCont1.SetFocus
        Exit Sub
    End If
    SBUCont2 = UCase(FN_TB_STRValidation(Len(txtCont2.Text), "SBU Contact Number 02", Trim(txtCont2.Text)))
    If Len(SBUCont2) = 0 Then
        txtCont2.SetFocus
        Exit Sub
    End If
    
    'Duplicate Checking
    STRSQL = "Select * from Inx_MSTR_SBU Where SBU_Code='" & SBUCode & "'"
    PR_SQL_Execution STRSQL
    If RS.RecordCount > 0 Then
        MsgBox "SBU Code Already Exist in the System", vbExclamation
        txtSBUCode.SetFocus
        Exit Sub
    End If
    
    STRSQL = "Select * from Inx_MSTR_SBU Where SBU_Code='" & SBUCode & "' and SBU_Name='" & Trim(txtSBUName.Text) & "'"
    PR_SQL_Execution STRSQL
    If RS.RecordCount > 0 Then
        MsgBox "Factory Name Already Exist in the System Under Entered SBU Code", vbExclamation
        txtComName.SetFocus
        Exit Sub
    End If
    
    STRSQL = "INSERT INTO Inx_MSTR_SBU(SBU_Code,Com_Code,SBU_Name,SBU_Add1,SBU_Add2,SBU_Cont1,SBU_Cont2) VALUES('" & SBUCode & "','" & ComCode & "','" & SBUName & "','" & SBUAdd1 & "','" & SBUAdd2 & "','" & SBUCont1 & "','" & SBUCont2 & "')"
    PR_SQL_Execution STRSQL
    PR_FMT_Grid
    Exit Sub
er_eh:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub cmdDelete_Click()
    If MsgBox("Are you sure Do you Want to delete ", vbQuestion + vbYesNo, "Delete Record") = vbYes Then
        If MSFSBU.Rows > -1 Then
            STRSQL = "Delete from Inx_MSTR_SBU where SBU_Code='" & Trim(MSFSBU.TextMatrix(MSFSBU.Row, 0)) & "'"
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
    If Trim(txtSBUName.Text) = "" Then
        STRSQL = "Select * from InxV_MSTR_SBU Order by SBU_Code"
    Else
        STRSQL = "Select * from InxV_MSTR_SBU Where SBU_Name like'" & Trim(txtSBUName.Text) & "%'"
    End If
    PR_SQL_Execution STRSQL
    With MSFSBU
        .Cols = 7
        .Rows = 1
        R = 1
        Do While RS.EOF = False
            .Rows = .Rows + 1
            .TextMatrix(R, 0) = Trim(RS!SBU_Code)
            .TextMatrix(R, 1) = Trim(RS!SBU_Name)
            .TextMatrix(R, 2) = Trim(RS!Com_Name)
            .TextMatrix(R, 3) = Trim(RS!SBU_Add1)
            .TextMatrix(R, 4) = Trim(RS!SBU_Add2)
            .TextMatrix(R, 5) = Trim(RS!SBU_Cont1)
            .TextMatrix(R, 6) = Trim(RS!SBU_Cont2)
            R = R + 1
            RS.MoveNext
        Loop
        .TextMatrix(0, 0) = "Code"
        .TextMatrix(0, 1) = "SBU Name"
        .TextMatrix(0, 2) = "Company Name"
        .TextMatrix(0, 3) = "Address Line 1"
        .TextMatrix(0, 4) = "Address Line 2"
        .TextMatrix(0, 5) = "Phone No.1"
        .TextMatrix(0, 6) = "Phone No.2"
        
        .ColWidth(0) = 800
        .ColWidth(1) = 2000
        .ColWidth(2) = 2000
        .ColWidth(3) = 2000
        .ColWidth(4) = 2000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
    End With
End Sub

Private Sub Form_Load()
    PR_FMT_Grid
    PR_Fill_Company
End Sub

Public Sub PR_Fill_Company()
    STRSQL = "Select * from Inx_MSTR_Company Order by Com_Name"
    PR_SQL_Execution STRSQL
    dcmbComName.ListField = "Com_Name"
    Set dcmbComName.RowSource = RS
End Sub
