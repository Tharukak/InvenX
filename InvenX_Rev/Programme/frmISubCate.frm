VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmISubCat 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Item Sub category Master"
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7545
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
   Picture         =   "frmISubCate.frx":0000
   ScaleHeight     =   8940
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid MSFSubCategory 
      Height          =   6855
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   12091
      _Version        =   393216
   End
   Begin VB.TextBox txtSubCatName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2280
      TabIndex        =   9
      Top             =   1320
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      Height          =   7695
      Left            =   6120
      TabIndex        =   4
      Top             =   840
      Width           =   1335
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H8000000A&
         Caption         =   "E&xit"
         Height          =   375
         Left            =   120
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   7200
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H8000000A&
         Caption         =   "&Add"
         Height          =   375
         Left            =   120
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
   End
   Begin MSDataListLib.DataCombo dcmbCategory 
      Height          =   315
      Left            =   2280
      TabIndex        =   2
      Top             =   960
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   "DataCombo1"
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   8565
      Width           =   7545
      _ExtentX        =   13309
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Sub Catagory Name"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Catagory Name"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Sub Category Configuration"
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
Attribute VB_Name = "frmISubCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
On Error GoTo er_EH:
    Dim SubCatName As String
    Dim CatName As String
    
    CatName = UCase(FN_TB_STRValidation(Len(dcmbCategory.Text), "Item Category", Trim(dcmbCategory.Text)))
    If Len(CatName) = 0 Then
        dcmbCategory.SetFocus
        Exit Sub
    End If

    SubCatName = UCase(FN_TB_STRValidation(Len(txtSubCatName.Text), "Item Sub Category", Trim(txtSubCatName.Text)))
    If Len(SubCatName) = 0 Then
        txtSubCatName.SetFocus
        Exit Sub
    End If
    
    'Duplicate Checking
    STRSQL = "Select * from InxV_MSTR_SubCategory Where Cat_Name='" & Trim(dcmbCategory.Text) & "' and SubCat_Name='" & SubCatName & "'"
    PR_SQL_Execution STRSQL
    
    If RS.RecordCount > 0 Then
        MsgBox "Sub Sub Category Already in the System", vbExclamation
    Else
        STRSQL = "Select Cat_ID from Inx_MSTR_Category Where Cat_Name='" & Trim(dcmbCategory.Text) & "'"
        PR_SQL_Execution STRSQL
        If RS.EOF = False Then
            Cat_ID = Val(RS!Cat_ID)
            STRSQL = "INSERT INTO Inx_MSTR_SubCategory(Cat_ID,SubCat_Name) VALUES(" & Cat_ID & ",'" & SubCatName & "')"
            PR_SQL_Execution STRSQL
            PR_FMT_Grid
        Else
            MsgBox "Category NOT Found", vbExclamation
            Exit Sub
        End If
    End If
    Exit Sub
er_EH:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub cmdDelete_Click()
    If MsgBox("Are you sure Do you Want to delete ", vbQuestion + vbYesNo, "Delete Record") = vbYes Then
        If MSFSubCategory.Rows > -1 Then
            STRSQL = "Delete from Inx_MSTR_SubCategory where SubCat_ID=" & Val(MSFSubCategory.TextMatrix(MSFSubCategory.Row, 0))
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
    If Trim(txtSubCatName.Text) = "" Then
        STRSQL = "Select * from InxV_MSTR_SubCategory Order by Cat_ID"
    Else
        STRSQL = "Select * from InxV_MSTR_SubCategory Where SubCat_Name like '" & Trim(txtSubCatName.Text) & "%'"
    End If
    PR_SQL_Execution STRSQL
    With MSFSubCategory
        .Cols = 3
        .Rows = 1
        R = 1
        Do While RS.EOF = False
            .Rows = .Rows + 1
            .TextMatrix(R, 0) = Trim(RS!SubCat_ID)
            .TextMatrix(R, 1) = Trim(RS!Cat_Name)
            .TextMatrix(R, 2) = Trim(RS!SubCat_Name)
            R = R + 1
            RS.MoveNext
        Loop
        .TextMatrix(0, 0) = "ID"
        .TextMatrix(0, 1) = "Category Name"
        .TextMatrix(0, 2) = "Sub Category Name"
        
        .ColWidth(0) = 800
        .ColWidth(1) = 2000
        .ColWidth(2) = 3000
    End With
End Sub

Private Sub Form_Load()
    PR_FMT_Grid
    PR_Fill_Category
End Sub
Public Sub PR_Fill_Category()
    STRSQL = "Select * from Inx_MSTR_Category Order by Cat_Name"
    PR_SQL_Execution STRSQL
    dcmbCategory.ListField = "Cat_Name"
    Set dcmbCategory.RowSource = RS
End Sub
