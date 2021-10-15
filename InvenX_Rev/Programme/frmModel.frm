VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmModel 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Model Master"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9135
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
   Picture         =   "frmModel.frx":0000
   ScaleHeight     =   9180
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtModelNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5160
      TabIndex        =   4
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox txtModelName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   1680
      Width           =   2655
   End
   Begin MSDataListLib.DataCombo dcmbCategory 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin VB.Frame Frame1 
      Height          =   6735
      Left            =   7680
      TabIndex        =   9
      Top             =   2040
      Width           =   1335
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H8000000A&
         Caption         =   "E&xit"
         Height          =   375
         Left            =   120
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6240
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H8000000A&
         Caption         =   "&Add"
         Height          =   375
         Left            =   120
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   5
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
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   8805
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   21167
            MinWidth        =   21167
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFModel 
      Height          =   6735
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   11880
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo dcmbBrand 
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Top             =   1320
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Model Number"
      Height          =   255
      Left            =   4080
      TabIndex        =   14
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Brand Name"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Model Name"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Model Configuration"
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
      TabIndex        =   0
      Top             =   480
      Width           =   5295
   End
End
Attribute VB_Name = "frmModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
On Error GoTo er_eh:
    Dim CatName, BrandName, ModelName, ModelNo As String
    Dim CatID, BrandID As Integer
    
    CatName = UCase(FN_TB_STRValidation(Len(dcmbCategory.Text), "Item Category", Trim(dcmbCategory.Text)))
    If Len(CatName) = 0 Then
        dcmbCategory.SetFocus
        Exit Sub
    End If
    BrandName = UCase(FN_TB_STRValidation(Len(dcmbBrand.Text), "Brand Name", Trim(dcmbBrand.Text)))
    If Len(CatName) = 0 Then
        dcmbBrand.SetFocus
        Exit Sub
    End If
    ModelName = UCase(FN_TB_STRValidation(Len(txtModelName.Text), "Model Name", Trim(txtModelName.Text)))
    If Len(ModelName) = 0 Then
        txtModelName.SetFocus
        Exit Sub
    End If
    ModelNo = UCase(FN_TB_STRValidation(Len(txtModelNo.Text), "Model Number", Trim(txtModelNo.Text)))
    If Len(ModelNo) = 0 Then
        txtModelNo.SetFocus
        Exit Sub
    End If
    
    STRSQL = "Select * from Inx_MSTR_Category Where Cat_Name='" & CatName & "'"
    PR_SQL_Execution STRSQL
    If RS.EOF = False Then
        CatID = Val(RS!Cat_ID)
    End If
    
    STRSQL = "Select * from Inx_MSTR_Brand Where Brand_Name='" & BrandName & "'"
    PR_SQL_Execution STRSQL
    If RS.EOF = False Then
        BrandID = Val(RS!Brand_ID)
    End If
    
    STRSQL = "INSERT INTO Inx_MSTR_Model(Cat_ID,Brand_ID,Model_Name,Model_No) Values(" & CatID & "," & BrandID & ",'" & ModelName & "','" & ModelNo & "')"
    PR_SQL_Execution STRSQL
    PR_FMT_Grid
    Exit Sub
    
er_eh:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub cmdDelete_Click()
    If MsgBox("Are you sure Do you Want to delete ", vbQuestion + vbYesNo, "Delete Record") = vbYes Then
        If MSFModel.Rows > -1 Then
            STRSQL = "Delete from Inx_MSTR_Model where Model_ID=" & Val(MSFModel.TextMatrix(MSFModel.Row, 0))
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

Private Sub Form_Load()
    PR_FillCat
    PR_FillBrand
    PR_FMT_Grid
End Sub

Public Sub PR_FillCat()
    STRSQL = "Select * from Inx_MSTR_Category Order by Cat_Name"
    PR_SQL_Execution STRSQL
    dcmbCategory.ListField = "Cat_Name"
    Set dcmbCategory.RowSource = RS
End Sub

Public Sub PR_FillBrand()
    STRSQL = "Select * from Inx_MSTR_Brand Order by Brand_Name"
    PR_SQL_Execution STRSQL
    dcmbBrand.ListField = "Brand_Name"
    Set dcmbBrand.RowSource = RS
End Sub

Public Sub PR_FMT_Grid()
    STRSQL = "Select * from InxV_MSTR_Model Order by Model_Name"
    PR_SQL_Execution STRSQL
    With MSFModel
    .Cols = 5
    .Rows = 1
    R = 1
    Do While RS.EOF = False
        .Rows = .Rows + 1
        .TextMatrix(R, 0) = Trim(RS!Model_ID)
        .TextMatrix(R, 1) = Trim(RS!Model_Name)
        .TextMatrix(R, 2) = Trim(RS!Model_No)
        .TextMatrix(R, 3) = Trim(RS!Cat_Name)
        .TextMatrix(R, 4) = Trim(RS!Brand_Name)
        R = R + 1
        RS.MoveNext
    Loop
    .TextMatrix(0, 0) = "ID"
    .TextMatrix(0, 1) = "Model Name"
    .TextMatrix(0, 2) = "Model No"
    .TextMatrix(0, 3) = "Category Name"
    .TextMatrix(0, 4) = "Brand Name"
    
    .ColWidth(0) = 800
    .ColWidth(1) = 2000
    .ColWidth(2) = 2000
    .ColWidth(3) = 2000
    .ColWidth(4) = 2000
    
    End With
End Sub

