VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmItemTrans 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Item Transaction"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15360
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
   Picture         =   "frmItemTrans.frx":0000
   ScaleHeight     =   8295
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H8000000A&
      Caption         =   "E&xit"
      Height          =   375
      Left            =   14040
      MaskColor       =   &H00404040&
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   6615
      Left            =   6480
      TabIndex        =   39
      Top             =   840
      Width           =   8775
      Begin VB.TextBox txtCUpdate 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   1920
         TabIndex        =   55
         Top             =   960
         Width           =   5415
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Process"
         Height          =   1695
         Left            =   7560
         TabIndex        =   54
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtRemarks 
         Height          =   285
         Left            =   1920
         TabIndex        =   53
         Top             =   1320
         Width           =   5415
      End
      Begin MSFlexGridLib.MSFlexGrid MSFHistory 
         Height          =   4455
         Left            =   120
         TabIndex        =   48
         Top             =   2040
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   7858
         _Version        =   393216
         Appearance      =   0
      End
      Begin MSDataListLib.DataCombo dcmbTMethod 
         Height          =   315
         Left            =   1920
         TabIndex        =   43
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbTo 
         Height          =   315
         Left            =   1920
         TabIndex        =   46
         Top             =   960
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         Text            =   ""
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Remarks"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Transaction History"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   1680
         Width           =   7215
      End
      Begin VB.Label lblFrom 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1920
         TabIndex        =   47
         Top             =   600
         Width           =   5415
      End
      Begin VB.Label Label43 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "To"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label42 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "From"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label41 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Transaction Method"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   8040
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   14
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "9/7/2021"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "1:47 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel12 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel13 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel14 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label lblItemID 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   51
      Top             =   840
      Width           =   4335
   End
   Begin VB.Label Label46 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Item ID"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   50
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Transaction"
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
      TabIndex        =   41
      Top             =   480
      Width           =   5295
   End
   Begin VB.Label lblRemarks 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   38
      Top             =   7680
      Width           =   4335
   End
   Begin VB.Label lblDOP 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   37
      Top             =   7320
      Width           =   4335
   End
   Begin VB.Label lblItemStat 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   36
      Top             =   6960
      Width           =   4335
   End
   Begin VB.Label lblPValue 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   35
      Top             =   6600
      Width           =   4335
   End
   Begin VB.Label lblPMethod 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   34
      Top             =   6240
      Width           =   4335
   End
   Begin VB.Label lblSup 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   33
      Top             =   5880
      Width           =   4335
   End
   Begin VB.Label lblCAPEX 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   32
      Top             =   5520
      Width           =   4335
   End
   Begin VB.Label lblFAR 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   31
      Top             =   5160
      Width           =   4335
   End
   Begin VB.Label lblSerial 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   30
      Top             =   4800
      Width           =   4335
   End
   Begin VB.Label lblModelNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   29
      Top             =   4440
      Width           =   4335
   End
   Begin VB.Label lblModelName 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   28
      Top             =   4080
      Width           =   4335
   End
   Begin VB.Label lblBrand 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   27
      Top             =   3720
      Width           =   4335
   End
   Begin VB.Label lblSubCat 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   26
      Top             =   3360
      Width           =   4335
   End
   Begin VB.Label lblCat 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   25
      Top             =   3000
      Width           =   4335
   End
   Begin VB.Label lblADUser 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   24
      Top             =   2640
      Width           =   4335
   End
   Begin VB.Label lblDivision 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   23
      Top             =   2280
      Width           =   4335
   End
   Begin VB.Label lblPlant 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   22
      Top             =   1920
      Width           =   4335
   End
   Begin VB.Label lblSBU 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   21
      Top             =   1560
      Width           =   4335
   End
   Begin VB.Label lblCompany 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   20
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Category Name"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sub Category Name"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Brand Name"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Model Name"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Model Number"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Company Name"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SBU Name"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Division Name"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AD User ID"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Supplier Name"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Purchase Method"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Purchase Value"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Item Status"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Remarks"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   7680
      Width           =   1695
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date of Purchase"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   7320
      Width           =   1695
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Seriall Number"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Plant Name"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FAR Tag"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label20 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CAPEX ID"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Width           =   1695
   End
End
Attribute VB_Name = "frmItemTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Selvar As String
Dim From_val As String
Dim To_Val As String
Private Sub cmdAdd_Click()
On Error GoTo er_Eh:
    Dim Trans_Nethod, To_Field As String
    Trans_Nethod = UCase(FN_TB_STRValidation(Len(dcmbTMethod.Text), "Transaction Method", Trim(dcmbTMethod.Text)))
    If dcmbTMethod.Text = "SERIAL NUMBER CHANGE" Or dcmbTMethod.Text = "FAR TAG CHANGE" Then
        To_Field = UCase(FN_TB_STRValidation(Len(txtCUpdate.Text), "Transaction Destination", Trim(txtCUpdate.Text)))
    Else
        To_Field = UCase(FN_TB_STRValidation(Len(dcmbTo.Text), "Transaction Destination", Trim(dcmbTo.Text)))
    End If
    
    Selvar = Trim(dcmbTMethod.Text)
    Set dcmbTo.RowSource = Nothing
    
    Select Case Selvar
        Case "DIVISION TRANSFER"
        STRSQL = "Update Inx_Items set Div_ID=(Select Div_ID from Inx_MSTR_Division Where Div_Name='" & Trim(dcmbTo.Text) & "') Where Item_ID=" & Val(lblItemID.Caption)
        PR_SQL_Execution STRSQL
        From_val = Trim(lblDivision.Caption)
        To_Val = Trim(dcmbTo.Text)
        
        Case "PLANT TRANSFER"
        STRSQL = "Update Inx_Items set Plant_Code=(Select Plant_Code from Inx_MSTR_Plant Where Plant_Name='" & Trim(dcmbTo.Text) & "') Where Item_ID=" & Val(lblItemID.Caption)
        PR_SQL_Execution STRSQL
        From_val = Trim(lblPlant.Caption)
        To_Val = Trim(dcmbTo.Text)
          
        Case "SBU TRANSFER"
        STRSQL = "Update Inx_Items set SBU_Code=(Select SBU_Code from Inx_MSTR_SBU Where SBU_Name='" & Trim(dcmbTo.Text) & "') Where Item_ID=" & Val(lblItemID.Caption)
        PR_SQL_Execution STRSQL
        From_val = Trim(lblSBU.Caption)
        To_Val = Trim(dcmbTo.Text)
          
        Case "ITEM CONDITION STATUS CHANGE"
        STRSQL = "Update Inx_Items set I_Stat_ID=(Select I_Stat_ID from Inx_MSTR_ItemStat Where I_Stat='" & Trim(dcmbTo.Text) & "') Where Item_ID=" & Val(lblItemID.Caption)
        PR_SQL_Execution STRSQL
        From_val = Trim(lblItemStat.Caption)
        To_Val = Trim(dcmbTo.Text)
          
        Case "USER TRANSFER"
        STRSQL = "Update Inx_Items set AD_UserID='" & Trim(dcmbTo.Text) & "' Where Item_ID=" & Val(lblItemID.Caption)
        PR_SQL_Execution STRSQL
        From_val = Trim(lblADUser.Caption)
        To_Val = Trim(dcmbTo.Text)
          
        Case "INHOUSE STATUS CHANGE"
        STRSQL = "Update Inx_Items set Inhoused=" & Val(dcmbTo.Text) & " Where Item_ID=" & Val(lblItemID.Caption)
        PR_SQL_Execution STRSQL
        From_val = Trim(lblItemStat.Caption)
        To_Val = Trim(dcmbTo.Text)
        
        Case "SERIAL NUMBER CHANGE"
        STRSQL = "Update Inx_Items set Serial_No='" & Trim(txtCUpdate.Text) & "' Where Item_ID='" & Val(lblItemID.Caption) & "'"
        PR_SQL_Execution STRSQL
        From_val = Trim(lblSerial.Caption)
        To_Val = Trim(txtCUpdate.Text)
        
        Case "FAR TAG CHANGE"
        STRSQL = "Update Inx_Items set FAR_Tag='" & Trim(txtCUpdate.Text) & "' Where Item_ID='" & Val(lblItemID.Caption) & "'"
        PR_SQL_Execution STRSQL
        From_val = Trim(lblFAR.Caption)
        To_Val = Trim(txtCUpdate.Text)
    End Select
    PR_History
    
    MsgBox Selvar & " UPDATED SUCCESSFULLY", vbInformation
    Exit Sub
    
er_Eh:
    MsgBox Err.Description, vbCritical
    Err.Clear
    
    
End Sub

Private Sub cmdExit_Click()
    Close All
    Unload Me
End Sub

Private Sub dcmbTMethod_Change()
    PR_Validation
End Sub

Private Sub Form_Load()
    STRSQL = "Select * from Inx_MSTR_Transaction Order by Trans_Name"
    PR_SQL_Execution STRSQL
    dcmbTMethod.ListField = "Trans_Name"
    Set dcmbTMethod.RowSource = RS
    PR_GRD_History
    dcmbTo.Visible = True
    txtCUpdate.Visible = False
End Sub

Private Sub lblItemID_Change()
    STRSQL = "Select * from InxV_Items Where Item_ID=" & Val(lblItemID.Caption)
    PR_SQL_Execution STRSQL
    If RS.EOF = False Then
        lblCompany.Caption = RS!Com_Name
        If IsNull(RS!SBU_Name) = False Then
            lblSBU.Caption = RS!SBU_Name
        Else
            lblSBU.Caption = ""
        End If
        If IsNull(RS!Plant_Name) = False Then
            lblPlant.Caption = RS!Plant_Name
        Else
            lblPlant.Caption = ""
        End If
        
        If IsNull(RS!Div_Name) = False Then
            lblDivision.Caption = RS!Div_Name
        Else
            lblDivision.Caption = ""
        End If
        lblADUser.Caption = RS!AD_UserID
        lblCat.Caption = RS!cat_Name
        lblSubCat.Caption = RS!SubCat_Name
        lblBrand.Caption = RS!Brand_Name
        lblModelName.Caption = RS!Model_Name
        lblModelNo.Caption = RS!Model_No
        lblSerial.Caption = RS!Serial_No
        lblFAR.Caption = RS!FAR_Tag
        lblCAPEX.Caption = RS!CAPEX_ID
        lblSup.Caption = "N/A"
        lblPMethod.Caption = RS!P_Method_Name
        lblPValue.Caption = RS!P_Value
        lblItemStat.Caption = RS!I_Stat
        lblDOP.Caption = RS!P_Date
        lblRemarks.Caption = RS!Remarks
        PR_GRD_History
    End If
End Sub

Public Sub PR_Validation()
    Dim Selvar As String
    Selvar = Trim(dcmbTMethod.Text)
    dcmbTo.Text = ""
    Set dcmbTo.RowSource = Nothing
    Select Case Selvar
        Case "DIVISION TRANSFER"
          STRSQL = "Select Div_Name from Inx_MSTR_Division"
          PR_SQL_Execution STRSQL
          If RS.EOF = False Then
                dcmbTo.ListField = "Div_Name"
                Set dcmbTo.RowSource = RS
          End If
          
          Case "PLANT TRANSFER"
          STRSQL = "Select Plant_Name from Inx_MSTR_Plant"
          PR_SQL_Execution STRSQL
          If RS.EOF = False Then
                dcmbTo.ListField = "Plant_Name"
                Set dcmbTo.RowSource = RS
          End If
          
          Case "SBU TRANSFER"
          STRSQL = "Select SBU_Name from Inx_MSTR_SBU"
          PR_SQL_Execution STRSQL
          If RS.EOF = False Then
                dcmbTo.ListField = "SBU_Name"
                Set dcmbTo.RowSource = RS
          End If
          
          Case "ITEM CONDITION STATUS CHANGE"
          STRSQL = "Select I_Stat from Inx_MSTR_ItemStat"
          PR_SQL_Execution STRSQL
          If RS.EOF = False Then
                dcmbTo.ListField = "I_Stat"
                Set dcmbTo.RowSource = RS
          End If
          
          Case "USER TRANSFER"
          STRSQL = "SELECT AD_UserID FROM Inx_MSTR_AD_Users Order by AD_UserID"
          PR_SQL_Execution STRSQL
          If RS.EOF = False Then
                dcmbTo.ListField = "AD_UserID"
                Set dcmbTo.RowSource = RS
          End If
          
          Case "INHOUSE STATUS CHANGE"
          Set RS = New ADODB.Recordset
          STRSQL = "Select Inhoused from Inx_Items Group by Inhoused"
          PR_SQL_Execution STRSQL
          If RS.EOF = False Then
                dcmbTo.ListField = "Inhoused"
                Set dcmbTo.RowSource = RS
          End If
          
          Case "SERIAL NUMBER CHANGE"
            MenuAccess 13
            If RightsMode = 1 Then
                dcmbTo.Visible = False
                txtCUpdate.Visible = True
            End If
            
          Case "FAR TAG CHANGE"
            MenuAccess 13
            If RightsMode = 1 Then
                dcmbTo.Visible = False
                txtCUpdate.Visible = True
            End If

    End Select
End Sub

Public Sub PR_History()
    STRSQL = "Insert Into Inx_Trans_History(Item_ID,Trans_Code,U_ID,Trans_Date_Time,From_Val,To_Val) Values(" & Val(lblItemID.Caption) & ",(Select Trans_Code from Inx_MSTR_Transaction Where Trans_Name='" & Selvar & "')," & UserID & ",'" & Date + Time & "','" & Trim(From_val) & "','" & Trim(To_Val) & "')"
    PR_SQL_Execution STRSQL
    PR_GRD_History
End Sub

Public Sub PR_GRD_History()
    MSFHistory.Cols = 3
    MSFHistory.TextMatrix(0, 0) = "TRANSACTION METHOD"
    MSFHistory.TextMatrix(0, 1) = "USER ID"
    MSFHistory.TextMatrix(0, 2) = "DATE TIME"
    
    STRSQL = "SELECT * FROM InxV_Trans_History WHERE ITEM_ID=" & Val(lblItemID.Caption) & "ORDER BY TRANS_DATE_TIME DESC"
    PR_SQL_Execution STRSQL
    R = 1
    Do While RS.EOF = False
        MSFHistory.Rows = R + 1
        MSFHistory.TextMatrix(R, 0) = Trim(RS!TRANS_NAME)
        MSFHistory.TextMatrix(R, 1) = Val(RS!U_ID)
        MSFHistory.TextMatrix(R, 2) = Format(RS!TRANS_DATE_TIME, "dd-MMM-yyyy HH:mm:ss")
        R = R + 1
        RS.MoveNext
    Loop
    
    MSFHistory.ColWidth(0) = 3000
    MSFHistory.ColWidth(1) = 1000
    MSFHistory.ColWidth(2) = 2500
End Sub

