VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmItem 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Item Creation"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17295
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
   Picture         =   "frmItem.frx":0000
   ScaleHeight     =   8730
   ScaleWidth      =   17295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCSV 
      Caption         =   "&CSV"
      Height          =   615
      Left            =   15240
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   7800
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   615
      Left            =   16080
      TabIndex        =   48
      Top             =   7800
      Width           =   1095
   End
   Begin VB.TextBox txtCAPEX 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   13
      Top             =   5160
      Width           =   3855
   End
   Begin VB.TextBox txtFAR 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   12
      Top             =   4800
      Width           =   3855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Searching"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      TabIndex        =   42
      Top             =   7680
      Width           =   9255
      Begin VB.CommandButton cmdSearch 
         Height          =   495
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   160
         Width           =   735
      End
      Begin VB.TextBox txtVar 
         Height          =   285
         Left            =   2280
         TabIndex        =   44
         Top             =   240
         Width           =   2895
      End
      Begin VB.ComboBox cmbFld 
         Height          =   315
         ItemData        =   "frmItem.frx":312D
         Left            =   120
         List            =   "frmItem.frx":3140
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.TextBox txtSerialNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   11
      Top             =   4440
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   38
      Top             =   7680
      Width           =   5655
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   2760
         TabIndex        =   39
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdde 
         Caption         =   "&Add"
         Height          =   375
         Left            =   4080
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFItems 
      Height          =   6855
      Left            =   5880
      TabIndex        =   36
      Top             =   840
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   12091
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker DTPDOP 
      Height          =   300
      Left            =   1920
      TabIndex        =   18
      Top             =   6960
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   103088131
      CurrentDate     =   44214
   End
   Begin VB.TextBox txtRemarks 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   19
      Top             =   7320
      Width           =   3855
   End
   Begin VB.TextBox txtPValue 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   16
      Top             =   6240
      Width           =   3855
   End
   Begin VB.TextBox txtADUserName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   2280
      Width           =   3855
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   8475
      Width           =   17295
      _ExtentX        =   30506
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "8/23/2021"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "8:55 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   24694
            MinWidth        =   24694
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo dcmbCompany 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcmbSBU 
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Top             =   1200
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcmbDivision 
      Height          =   315
      Left            =   1920
      TabIndex        =   4
      Top             =   1920
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcmbCat 
      Height          =   315
      Left            =   1920
      TabIndex        =   6
      Top             =   2640
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcmbSubCat 
      Height          =   315
      Left            =   1920
      TabIndex        =   7
      Top             =   3000
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcmbBrand 
      Height          =   315
      Left            =   1920
      TabIndex        =   8
      Top             =   3360
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcmbModel 
      Height          =   315
      Left            =   1920
      TabIndex        =   9
      Top             =   3720
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcmbModelNo 
      Height          =   315
      Left            =   1920
      TabIndex        =   10
      Top             =   4080
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcmbSupplier 
      Height          =   315
      Left            =   1920
      TabIndex        =   14
      Top             =   5520
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcmbPMethod 
      Height          =   315
      Left            =   1920
      TabIndex        =   15
      Top             =   5880
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcmbItemStat 
      Height          =   315
      Left            =   1920
      TabIndex        =   17
      Top             =   6600
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcmbPlant 
      Height          =   315
      Left            =   1920
      TabIndex        =   3
      Top             =   1560
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
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
      TabIndex        =   47
      Top             =   5160
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
      TabIndex        =   46
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
      TabIndex        =   41
      Top             =   1560
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
      TabIndex        =   40
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Creation"
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
      TabIndex        =   37
      Top             =   480
      Width           =   5295
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
      TabIndex        =   35
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
      TabIndex        =   34
      Top             =   7320
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
      TabIndex        =   33
      Top             =   6600
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
      TabIndex        =   32
      Top             =   6240
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
      TabIndex        =   31
      Top             =   5880
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
      TabIndex        =   30
      Top             =   5520
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
      TabIndex        =   29
      Top             =   2280
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
      TabIndex        =   28
      Top             =   1920
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
      TabIndex        =   27
      Top             =   1200
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
      TabIndex        =   26
      Top             =   840
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
      TabIndex        =   25
      Top             =   4080
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
      TabIndex        =   24
      Top             =   3720
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
      TabIndex        =   23
      Top             =   3360
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
      TabIndex        =   22
      Top             =   3000
      Width           =   1695
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
      TabIndex        =   21
      Top             =   2640
      Width           =   1695
   End
End
Attribute VB_Name = "frmItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Tbl_Name As String
Dim Sort_Name As String
Dim Fld_Name As String

Private Sub cmdAdde_Click()
On Error GoTo er_Eh:
    Dim FAR_Tag, CAPEX_ID, Company, Company_Code, SBU, SBU_Code, Plant, Plant_Code, Division, Category, SubCategory, Brand, Model, ModelNo, Supplier, PMethod, ADUser, ItemStat, SerialNo, Remarks As String
    Dim Div_ID, Cat_ID, SubCat_ID, Brand_ID, Model_ID, Sup_ID, P_Method_ID, P_Value, Stat_ID As Long
    
    Company = UCase(FN_TB_STRValidation(Len(dcmbCompany.Text), "Company Name", Trim(dcmbCompany.Text)))
    If Len(Company) = 0 Then
        dcmbCompany.SetFocus
        Exit Sub
    Else
        STRSQL = "Select * from Inx_MSTR_Company Where Com_Name='" & Company & "'"
        PR_SQL_Execution STRSQL
        If RS.EOF = False Then
            Company_Code = Trim(RS!Com_Code)
        End If
    End If
    
    SBU = UCase(FN_TB_STRValidation(Len(dcmbSBU.Text), "SBU Name", Trim(dcmbSBU.Text)))
    If Len(SBU) = 0 Then
        dcmbSBU.SetFocus
        Exit Sub
    Else
        STRSQL = "Select * from Inx_MSTR_sbu Where Com_Code='" & Company_Code & "' and SBU_Name='" & SBU & "'"
        PR_SQL_Execution STRSQL
        If RS.EOF = False Then
            SBU_Code = Trim(RS!SBU_Code)
        End If
    End If
    
    Plant = UCase(FN_TB_STRValidation(Len(dcmbPlant.Text), "Plant Name", Trim(dcmbPlant.Text)))
    If Len(Plant) = 0 Then
        dcmbPlant.SetFocus
        Exit Sub
    Else
        STRSQL = "Select * from Inx_MSTR_Plant Where Plant_Name='" & Plant & "' and Com_Code='" & Company_Code & "'"
        PR_SQL_Execution STRSQL
        If RS.EOF = False Then
            Plant_Code = Trim(RS!Plant_Code)
        End If
    End If
    
    Division = UCase(FN_TB_STRValidation(Len(dcmbDivision.Text), "Division Name", Trim(dcmbDivision.Text)))
    If Len(Division) = 0 Then
        dcmbDivision.SetFocus
        Exit Sub
    Else
        STRSQL = "Select * from Inx_MSTR_Division Where Div_Name='" & Division & "'"
        PR_SQL_Execution STRSQL
        If RS.EOF = False Then
            Div_ID = Val(RS!Div_ID)
        End If
    End If
    
    ADUser = UCase(FN_TB_STRValidation(Len(txtADUserName.Text), "Active Directory User Name", Trim(txtADUserName.Text)))
    If Len(ADUser) = 0 Then
        txtADUserName.SetFocus
        Exit Sub
    End If
    
    Category = UCase(FN_TB_STRValidation(Len(dcmbCat.Text), "Category Name", Trim(dcmbCat.Text)))
    If Len(Category) = 0 Then
        dcmbCat.SetFocus
        Exit Sub
    Else
        STRSQL = "Select * from Inx_MSTR_Category Where Cat_Name='" & Category & "'"
        PR_SQL_Execution STRSQL
        If RS.EOF = False Then
            Cat_ID = Val(RS!Cat_ID)
        End If
    End If
    
    SubCategory = UCase(FN_TB_STRValidation(Len(dcmbSubCat.Text), "Sub Category Name", Trim(dcmbSubCat.Text)))
    If Len(SubCategory) = 0 Then
        dcmbSubCat.SetFocus
        Exit Sub
    Else
        STRSQL = "Select * from Inx_MSTR_SubCategory Where SubCat_Name='" & SubCategory & "' and Cat_ID=" & Cat_ID
        PR_SQL_Execution STRSQL
        If RS.EOF = False Then
            SubCat_ID = Val(RS!SubCat_ID)
        End If
    End If
    
    Brand = UCase(FN_TB_STRValidation(Len(dcmbBrand.Text), "Brand Name", Trim(dcmbBrand.Text)))
    If Len(Brand) = 0 Then
        dcmbBrand.SetFocus
        Exit Sub
    Else
        STRSQL = "Select * from Inx_MSTR_Brand Where Brand_Name='" & Brand & "'"
        PR_SQL_Execution STRSQL
        If RS.EOF = False Then
            Brand_ID = Val(RS!Brand_ID)
        End If
    End If
    
    Model = UCase(FN_TB_STRValidation(Len(dcmbModel.Text), "Model Name", Trim(dcmbModel.Text)))
    If Len(Model) = 0 Then
        dcmbModel.SetFocus
        Exit Sub
    End If
    
    ModelNo = UCase(FN_TB_STRValidation(Len(dcmbModelNo.Text), "Model Number", Trim(dcmbModelNo.Text)))
    If Len(ModelNo) = 0 Then
        dcmbModelNo.SetFocus
        Exit Sub
    Else
        STRSQL = "Select * from Inx_MSTR_Model Where Model_Name='" & Model & "' and Model_No='" & ModelNo & "'"
        PR_SQL_Execution STRSQL
        If RS.EOF = False Then
            Model_ID = Val(RS!Model_ID)
        End If
    End If
    
    Supplier = UCase(FN_TB_STRValidation(Len(dcmbSupplier.Text), "Supplier Name", Trim(dcmbSupplier.Text)))
    If Len(Supplier) = 0 Then
        dcmbSupplier.SetFocus
        Exit Sub
    Else
        STRSQL = "Select * from Inx_MSTR_Supplier Where Sup_Name='" & Supplier & "'"
        PR_SQL_Execution STRSQL
        If RS.EOF = False Then
            Sup_ID = Val(RS!Sup_ID)
        End If
    End If
    
    PMethod = UCase(FN_TB_STRValidation(Len(dcmbPMethod.Text), "Purchase Method", Trim(dcmbPMethod.Text)))
    If Len(PMethod) = 0 Then
        dcmbPMethod.SetFocus
        Exit Sub
    Else
        STRSQL = "Select * from Inx_MSTR_PMethods Where P_Method_Name='" & PMethod & "'"
        PR_SQL_Execution STRSQL
        If RS.EOF = False Then
            P_Method_ID = Val(RS!P_Method_ID)
        End If
    End If
    
    ItemStat = UCase(FN_TB_STRValidation(Len(dcmbItemStat.Text), "Item Status", Trim(dcmbItemStat.Text)))
    If Len(ItemStat) = 0 Then
        dcmbItemStat.SetFocus
        Exit Sub
    End If
    
    P_Value = UCase(FN_TB_STRValidation(Len(txtPValue.Text), "Purchase Value", Trim(txtPValue.Text)))
    If Len(P_Value) = 0 Then
        txtPValue.SetFocus
        Exit Sub
    End If
    
    ItemStat = UCase(FN_TB_STRValidation(Len(dcmbItemStat.Text), "Item Status", Trim(dcmbItemStat.Text)))
    If Len(ItemStat) = 0 Then
        dcmbItemStat.SetFocus
        Exit Sub
    Else
        STRSQL = "Select * from Inx_MSTR_ItemStat Where I_Stat='" & ItemStat & "'"
        PR_SQL_Execution STRSQL
        If RS.EOF = False Then
            Stat_ID = Val(RS!I_Stat_ID)
        End If
    End If
    
    SerialNo = UCase(FN_TB_STRValidation(Len(txtSerialNo.Text), "Serial Number", Trim(txtSerialNo.Text)))
    If Len(SerialNo) = 0 Then
        txtSerialNo.SetFocus
        Exit Sub
    End If
    
    FAR_Tag = UCase(FN_TB_STRValidation(Len(txtFAR.Text), "FAR Tag", Trim(txtFAR.Text)))
    If Len(FAR_Tag) = 0 Then
        txtFAR.SetFocus
        Exit Sub
    End If
    
    CAPEX_ID = UCase(FN_TB_STRValidation(Len(txtCAPEX.Text), "CAPEX ID", Trim(txtCAPEX.Text)))
    If Len(CAPEX_ID) = 0 Then
        txtCAPEX.SetFocus
        Exit Sub
    End If
    
    STRSQL = "INSERT INTO Inx_Items(Serial_No,FAR_Tag,CAPEX_ID,Com_Code,SBU_Code,Plant_Code,Div_ID,AD_UserID,Cat_ID,SubCat_ID,Brand_ID,Model_ID,Sup_ID,P_Method_ID,P_Value,P_Date,Remarks,I_Stat_ID,Inhoused)" _
    + " VALUES('" & SerialNo & "','" & FAR_Tag & "','" & CAPEX_ID & "','" & Company_Code & "','" & SBU_Code & "','" & Plant_Code & "','" & Div_ID & "','" & ADUser & "','" & Cat_ID & "','" & SubCat_ID & "','" & Brand_ID & "','" & Model_ID & "','" & Sup_ID & "','" & P_Method_ID & "','" & P_Value & "','" & DTPDOP.Value & "','" & Trim(txtRemarks.Text) & "','" & Stat_ID & "',1)"
    PR_SQL_Execution STRSQL
    
    STRSQL = "Select * from InxV_Items Order By Item_ID"
    FMT_Grid
    PR_Highlight
    Exit Sub
    
er_Eh:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub CMDcSV_Click()
On Error GoTo er_Eh:
    PR_Open_Con
    Dim PrintText As String
    Set RS = New ADODB.Recordset
    Open "C:\InvenX_Rev\CSV\" + Trim(Format(Date, "ddMMyy")) + ".csv" For Output As #1
    RS.Open "Select * from InxV_Items Order By Item_ID", Con, adOpenStatic, adLockReadOnly
    Dim RN As Long
    RN = 0
    PrintText = "Item_ID,Serial_No,FAR_Tag,CAPEX_ID,Com_Code,Com_Name,SBU_Code" _
      + ",SBU_Name,Plant_Code,Plant_Name,Cat_Name,SubCat_ID,SubCat_Name,Brand_Name" _
      + ",Model_Name,Model_No,P_Method_Name,Div_Name,AD_UserID,I_Stat"
    Print #1, PrintText
    Do While RS.EOF = False
         PrintText = Trim(Str(RS!Item_ID)) + "," + Trim(RS!Serial_No) + "," + Trim(RS!FAR_Tag) + "," + Trim(RS!CAPEX_ID) + "," + RS!Com_Code + "," + RS!Com_Name + "," + RS!SBU_Code + "," + RS!SBU_Name + "," + RS!Plant_Code + "," + RS!Plant_Name + "," + RS!cat_Name + "," + RS!SubCat_Name + "," + RS!Brand_Name + "," + RS!Model_Name + "," + RS!Model_No + "," + RS!P_Method_Name + "," + RS!Div_Name + "," + RS!AD_UserID + "," + RS!I_Stat
        
        Print #1, PrintText
        RS.MoveNext
    Loop
    PrintText = ""
    Close #1
    MsgBox "CSV Created Successfully", vbInformation
Exit Sub
er_Eh:
    Close #1
    MsgBox Err.Description
    Err.Clear
End Sub

Private Sub cmdExit_Click()
    Close All
    Unload Me
End Sub

Private Sub cmdSearch_Click()
On Error GoTo er_Eh:
    STRSQL = "Select * from InxV_Items Where " & Trim(cmbFld.Text) & " like '%" & Trim(txtVar.Text) & "%' Order By Item_ID"
    FMT_Grid
    Exit Sub
er_Eh:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub dcmbBrand_Change()
    STRSQL = "Select Model_Name from InxV_MSTR_Model Where Cat_Name='" & Trim(dcmbCat.Text) & "' and Brand_Name='" & Trim(dcmbBrand.Text) & "' Group by Model_Name Order by Model_Name"
    PR_SQL_Execution STRSQL
    dcmbModel.ListField = "Model_Name"
    Set dcmbModel.RowSource = RS
    If RS.RecordCount = 1 Then
        RS.MoveFirst
        dcmbModel.Text = Trim(RS!Model_Name)
    Else
        dcmbModel.ListField = "Model_Name"
        Set dcmbModel.RowSource = RS
    End If
End Sub

Private Sub dcmbCat_Change()
    STRSQL = "Select * from InxV_MSTR_SubCategory Where Cat_Name='" & Trim(dcmbCat.Text) & "' Order by SubCat_Name"
    PR_SQL_Execution STRSQL
    dcmbSubCat.ListField = "SubCat_Name"
    Set dcmbSubCat.RowSource = RS
End Sub

Private Sub dcmbCompany_Change()
    STRSQL = "Select * from InxV_MSTR_SBU Where Com_Name='" & Trim(dcmbCompany.Text) & "' Order by SBU_Name"
    PR_SQL_Execution STRSQL
    dcmbSBU.ListField = "SBU_Name"
    Set dcmbSBU.RowSource = RS
End Sub

Private Sub dcmbModel_Change()
    STRSQL = "Select Model_No from InxV_MSTR_Model Where Cat_Name='" & Trim(dcmbCat.Text) & "' and Brand_Name='" & Trim(dcmbBrand.Text) & "' and Model_Name='" & Trim(dcmbModel.Text) & "' Group by Model_No Order by Model_No"
    PR_SQL_Execution STRSQL
    dcmbModelNo.ListField = "Model_No"
    Set dcmbModelNo.RowSource = RS
    If RS.RecordCount = 1 Then
        RS.MoveFirst
        dcmbModelNo.Text = Trim(RS!Model_No)
    Else
        dcmbModelNo.ListField = "Model_NO"
        Set dcmbModelNo.RowSource = RS
    End If
End Sub

Private Sub dcmbSBU_Change()
    STRSQL = "Select * from InxV_MSTR_Plant Where Com_Name='" & Trim(dcmbCompany.Text) & "' and SBU_Name='" & Trim(dcmbSBU.Text) & "' Order by SBU_Name"
    PR_SQL_Execution STRSQL
    dcmbPlant.ListField = "Plant_Name"
    Set dcmbPlant.RowSource = RS
End Sub

Private Sub Form_Load()
    Dim ComName As String
    DTPDOP.Value = Date

    'Fill Company Name ------------
    Fld_Name = "com_Name"
    Tbl_Name = "Inx_MSTR_Company"
    Sort_Name = "Com_Name"
    X = FN_Fill_Combo(Fld_Name, Tbl_Name, Sort_Name)
    dcmbCompany.ListField = Fld_Name
    Set dcmbCompany.RowSource = RS
    If RS.RecordCount = 1 Then
        RS.MoveFirst
        dcmbCompany.Text = Trim(RS!Com_Name)
    End If
    
    'Fill Plant Name -------------
    Fld_Name = "Plant_Name"
    Tbl_Name = "Inx_MSTR_Plant"
    Sort_Name = "Plant_Name"
    X = FN_Fill_Combo(Fld_Name, Tbl_Name, Sort_Name)
    dcmbPlant.ListField = Fld_Name
    Set dcmbPlant.RowSource = RS
    If RS.RecordCount = 1 Then
        RS.MoveFirst
        dcmbPlant.Text = Trim(RS!Plant_Name)
    End If
    
    'Fill Diviaion Name -------------
    Fld_Name = "Div_Name"
    Tbl_Name = "Inx_MSTR_Division"
    Sort_Name = "Div_Name"
    X = FN_Fill_Combo(Fld_Name, Tbl_Name, Sort_Name)
    dcmbDivision.ListField = Fld_Name
    Set dcmbDivision.RowSource = RS
    If RS.RecordCount = 1 Then
        RS.MoveFirst
        dcmbCompany.Text = Trim(RS!Div_Name)
    End If
    
    'Fill Category Name -------------
    Fld_Name = "Cat_Name"
    Tbl_Name = "Inx_MSTR_Category"
    Sort_Name = "Cat_Name"
    X = FN_Fill_Combo(Fld_Name, Tbl_Name, Sort_Name)
    dcmbCat.ListField = Fld_Name
    Set dcmbCat.RowSource = RS
    If RS.RecordCount = 1 Then
        RS.MoveFirst
        dcmbCat.Text = Trim(RS!cat_Name)
    End If
    
    'Fill Sub Category Name -------------
    Fld_Name = "SubCat_Name"
    Tbl_Name = "Inx_MSTR_SubCategory"
    Sort_Name = "SubCat_Name"
    X = FN_Fill_Combo(Fld_Name, Tbl_Name, Sort_Name)
    dcmbSubCat.ListField = Fld_Name
    Set dcmbSubCat.RowSource = RS
    If RS.RecordCount = 1 Then
        RS.MoveFirst
        dcmbSubCat.Text = Trim(RS!SubCat_Name)
    End If
    
    'Fill Brand Name -------------
    Fld_Name = "Brand_Name"
    Tbl_Name = "Inx_MSTR_Brand"
    Sort_Name = "Brand_Name"
    X = FN_Fill_Combo(Fld_Name, Tbl_Name, Sort_Name)
    dcmbBrand.ListField = Fld_Name
    Set dcmbBrand.RowSource = RS
    If RS.RecordCount = 1 Then
        RS.MoveFirst
        dcmbBrand.Text = Trim(RS!Brand_Name)
    End If
    
    'Fill Supplier Name -------------
    Fld_Name = "Sup_Name"
    Tbl_Name = "Inx_MSTR_Supplier"
    Sort_Name = "Sup_Name"
    X = FN_Fill_Combo(Fld_Name, Tbl_Name, Sort_Name)
    dcmbSupplier.ListField = Fld_Name
    Set dcmbSupplier.RowSource = RS
    If RS.RecordCount = 1 Then
        RS.MoveFirst
        dcmbSupplier.Text = Trim(RS!Sup_Name)
    End If
    
    'FillPurchase Method------------
    Fld_Name = "P_Method_Name"
    Tbl_Name = "Inx_MSTR_PMethods"
    Sort_Name = "P_Method_Name"
    X = FN_Fill_Combo(Fld_Name, Tbl_Name, Sort_Name)
    dcmbPMethod.ListField = Fld_Name
    Set dcmbPMethod.RowSource = RS
    If RS.RecordCount = 1 Then
        RS.MoveFirst
        dcmbPMethod.Text = Trim(RS!P_Method_Name)
    End If
    
    
    'Fill Item Condition ------------
    Fld_Name = "I_Stat"
    Tbl_Name = "Inx_MSTR_ItemStat"
    Sort_Name = "I_Stat"
    X = FN_Fill_Combo(Fld_Name, Tbl_Name, Sort_Name)
    dcmbItemStat.ListField = Fld_Name
    Set dcmbItemStat.RowSource = RS
    If RS.RecordCount = 1 Then
        RS.MoveFirst
        dcmbItemStat.Text = Trim(RS!I_Stat)
    End If
    
    STRSQL = "Select * from InxV_Items Order By Item_ID"
    FMT_Grid
    cmbFld.Text = "SERIAL_NO"
    PR_Highlight
End Sub

Public Function FN_Fill_Combo(Fld_Name As String, Tbl_Name As String, Sort_Name As String)
    STRSQL = "Select " & Fld_Name & " from " & Tbl_Name & " Order by " & Sort_Name
    PR_SQL_Execution STRSQL
End Function

Public Function FN_Fill_Model_Name(Fld_Name As String, Tbl_Name As String, Sort_Name As String)
    STRSQL = "Select " & Fld_Name & " from " & Tbl_Name & " Order by " & Sort_Name
    PR_SQL_Execution STRSQL
End Function

Public Sub FMT_Grid()
    PR_SQL_Execution STRSQL
    MSFItems.Rows = 1
    MSFItems.Cols = 9
    R = 1
    Do While RS.EOF = False
        MSFItems.Rows = R + 1
        MSFItems.TextMatrix(R, 0) = Trim(RS!Item_ID)
        MSFItems.TextMatrix(R, 1) = Trim(RS!cat_Name)
        MSFItems.TextMatrix(R, 2) = Trim(RS!SubCat_Name)
        MSFItems.TextMatrix(R, 3) = Trim(RS!Brand_Name)
        MSFItems.TextMatrix(R, 4) = Trim(RS!Model_Name)
        MSFItems.TextMatrix(R, 5) = Trim(RS!Model_No)
        MSFItems.TextMatrix(R, 6) = Trim(RS!Serial_No)
        MSFItems.TextMatrix(R, 7) = Trim(RS!FAR_Tag)
        MSFItems.TextMatrix(R, 8) = Trim(RS!CAPEX_ID)
        
        'Highlight Not Inhoused Items
        MSFItems.Col = 0
        MSFItems.Row = R
        If RS!Inhoused = False Then
            MSFItems.CellBackColor = &HC0C0FF
        Else
            MSFItems.CellBackColor = &HFFFFFF
        End If
        
        R = R + 1
        RS.MoveNext
    Loop
    
    MSFItems.TextMatrix(0, 0) = "ITEM ID."
    MSFItems.TextMatrix(0, 1) = "CATEGORY NAME"
    MSFItems.TextMatrix(0, 2) = "SUB CATEGORY NAME"
    MSFItems.TextMatrix(0, 3) = "BRAND NAME"
    MSFItems.TextMatrix(0, 4) = "MODEL NAME"
    MSFItems.TextMatrix(0, 5) = "MODEL NO."
    MSFItems.TextMatrix(0, 6) = "SERIAL NO."
    MSFItems.TextMatrix(0, 7) = "FAR. TAG"
    MSFItems.TextMatrix(0, 8) = "CAPEX ID."
    
    MSFItems.ColWidth(0) = 800
    MSFItems.ColWidth(1) = 1800
    MSFItems.ColWidth(2) = 1800
    MSFItems.ColWidth(3) = 1500
    MSFItems.ColWidth(4) = 1500
    MSFItems.ColWidth(5) = 1500
    MSFItems.ColWidth(6) = 1500
    MSFItems.ColWidth(7) = 1500
    MSFItems.ColWidth(8) = 2500
    
End Sub

Public Sub PR_Highlight()
    For R = 1 To MSFItems.Rows - 1
        'For C = 1 To MSFItems.Cols - 1
            MSFItems.Col = 1
            MSFItems.Row = R
            MSFItems.CellBackColor = &H80000005
        'Next C
    Next R
End Sub

