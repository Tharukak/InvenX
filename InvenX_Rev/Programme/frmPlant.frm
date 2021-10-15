VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPlant 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   10065
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7695
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
   Picture         =   "frmPlant.frx":0000
   ScaleHeight     =   10065
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   6615
      Left            =   6240
      TabIndex        =   7
      Top             =   3000
      Width           =   1335
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
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H8000000A&
         Caption         =   "E&xit"
         Height          =   375
         Left            =   120
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   6120
         Width           =   1095
      End
   End
   Begin VB.TextBox txtPlantCode 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6480
      TabIndex        =   6
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtCont2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4560
      TabIndex        =   5
      Top             =   2640
      Width           =   2895
   End
   Begin VB.TextBox txtCont1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox txtAdd2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Top             =   2280
      Width           =   6015
   End
   Begin VB.TextBox txtAdd1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   1920
      Width           =   6015
   End
   Begin VB.TextBox txtPlantName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   1560
      Width           =   3615
   End
   Begin MSDataListLib.DataCombo dcmbComName 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   840
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin MSFlexGridLib.MSFlexGrid MSFPlant 
      Height          =   6495
      Left            =   120
      TabIndex        =   11
      Top             =   3000
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
      Top             =   9690
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
   Begin MSDataListLib.DataCombo dcmbSBUName 
      Height          =   315
      Left            =   1440
      TabIndex        =   24
      Top             =   1200
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label lblSBUCode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6480
      TabIndex        =   26
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "SBU Code"
      Height          =   255
      Left            =   5280
      TabIndex        =   25
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Plant Name"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Plant Configuration"
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
      Left            =   1320
      TabIndex        =   22
      Top             =   480
      Width           =   5295
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Plant Code"
      Height          =   255
      Left            =   5280
      TabIndex        =   21
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "2."
      Height          =   255
      Left            =   4200
      TabIndex        =   20
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No. 1"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Address Line 2"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Address Line 1"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SBU Name"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Code"
      Height          =   255
      Left            =   5280
      TabIndex        =   14
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblComCode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6480
      TabIndex        =   13
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "frmPlant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
On Error GoTo er_eh:
    Dim SBUCode, SBUName, PlantCode, PlantName, ComCode, ComName, PlantAdd1, PlantAdd2, PlantCont1, PlantCont2 As String
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
    
    SBUName = UCase(FN_TB_STRValidation(Len(dcmbSBUName.Text), "SBU Name", Trim(dcmbSBUName.Text)))
    If Len(SBUName) = 0 Then
        dcmbSBUName.SetFocus
        Exit Sub
    End If
    
    STRSQL = "Select * from Inx_MSTR_SBU Where SBU_Name='" & SBUName & "'"
    PR_SQL_Execution STRSQL
    If RS.EOF = False Then
        SBUCode = Trim(RS!SBU_Code)
    Else
        MsgBox "SBU NOT Found", vbExclamation
        Exit Sub
    End If
        
    PlantName = UCase(FN_TB_STRValidation(Len(txtPlantName.Text), "Plant Name", Trim(txtPlantName.Text)))
    If Len(PlantName) = 0 Then
        txtPlantName.SetFocus
        Exit Sub
    End If
    
    PlantCode = UCase(FN_TB_STRValidation(Len(txtPlantCode.Text), "Plant Code", Trim(txtPlantCode.Text)))
    If Len(PlantCode) = 0 Then
        txtPlantCode.SetFocus
        Exit Sub
    End If
    
    PlantAdd1 = UCase(FN_TB_STRValidation(Len(txtAdd1.Text), "Plant Address Line 01", Trim(txtAdd1.Text)))
    If Len(PlantAdd1) = 0 Then
        txtAdd1.SetFocus
        Exit Sub
    End If
    PlantAdd2 = UCase(FN_TB_STRValidation(Len(txtAdd2.Text), "Plant Address Line 02", Trim(txtAdd2.Text)))
    If Len(PlantAdd2) = 0 Then
        txtAdd2.SetFocus
        Exit Sub
    End If
    PlantCont1 = UCase(FN_TB_STRValidation(Len(txtCont1.Text), "Plant Contact Number 01", Trim(txtCont1.Text)))
    If Len(PlantCont1) = 0 Then
        txtCont1.SetFocus
        Exit Sub
    End If
    PlantCont2 = UCase(FN_TB_STRValidation(Len(txtCont2.Text), "Plant Contact Number 02", Trim(txtCont2.Text)))
    If Len(PlantCont2) = 0 Then
        txtCont2.SetFocus
        Exit Sub
    End If
    
    'Duplicate Checking
    STRSQL = "Select * from Inx_MSTR_Plant Where Plant_Code='" & PlantCode & "'"
    PR_SQL_Execution STRSQL
    If RS.RecordCount > 0 Then
        MsgBox "Plant Code Already Exist in the System", vbExclamation
        txtPlantCode.SetFocus
        Exit Sub
    End If
    
    STRSQL = "Select * from Inx_MSTR_Plant Where Plant_Code='" & PlantCode & "' and Plant_Name='" & Trim(txtPlantName.Text) & "'"
    PR_SQL_Execution STRSQL
    If RS.RecordCount > 0 Then
        MsgBox "Factory Name Already Exist in the System Under Entered Plant Code", vbExclamation
        txtComName.SetFocus
        Exit Sub
    End If
    
    STRSQL = "INSERT INTO Inx_MSTR_Plant(Plant_Code,Com_Code,SBU_Code,Plant_Name,Plant_Add1,Plant_Add2,Plant_Cont1,Plant_Cont2) VALUES('" & PlantCode & "','" & ComCode & "','" & SBUCode & "','" & PlantName & "','" & PlantAdd1 & "','" & PlantAdd2 & "','" & PlantCont1 & "','" & PlantCont2 & "')"
    PR_SQL_Execution STRSQL
    PR_FMT_Grid
    Exit Sub
er_eh:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub cmdDelete_Click()
    If MsgBox("Are you sure Do you Want to delete ", vbQuestion + vbYesNo, "Delete Record") = vbYes Then
        If MSFPlant.Rows > -1 Then
            STRSQL = "Delete from Inx_MSTR_Plant where Plant_Code='" & Trim(MSFPlant.TextMatrix(MSFPlant.Row, 0)) & "'"
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
    If Trim(txtPlantName.Text) = "" Then
        STRSQL = "Select * from InxV_MSTR_Plant Order by Plant_Code"
    Else
        STRSQL = "Select * from InxV_MSTR_Plant Where Plant_Name like'" & Trim(txtPlantName.Text) & "%'"
    End If
    PR_SQL_Execution STRSQL
    With MSFPlant
        .Cols = 7
        .Rows = 1
        R = 1
        Do While RS.EOF = False
            .Rows = .Rows + 1
            .TextMatrix(R, 0) = Trim(RS!Plant_Code)
            .TextMatrix(R, 1) = Trim(RS!Plant_Name)
            .TextMatrix(R, 2) = Trim(RS!Com_Name)
            .TextMatrix(R, 3) = Trim(RS!Plant_Add1)
            .TextMatrix(R, 4) = Trim(RS!Plant_Add2)
            .TextMatrix(R, 5) = Trim(RS!Plant_Cont1)
            .TextMatrix(R, 6) = Trim(RS!Plant_Cont2)
            R = R + 1
            RS.MoveNext
        Loop
        .TextMatrix(0, 0) = "Code"
        .TextMatrix(0, 1) = "Plant Name"
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
    
    STRSQL = "Select * from Inx_MSTR_SBU Order by SBU_Name"
    PR_SQL_Execution STRSQL
    dcmbSBUName.ListField = "SBU_Name"
    Set dcmbSBUName.RowSource = RS
End Sub

