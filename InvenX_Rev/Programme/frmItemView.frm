VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmItemView 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inem View"
   ClientHeight    =   10455
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14910
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
   Picture         =   "frmItemView.frx":0000
   ScaleHeight     =   10455
   ScaleWidth      =   14910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   360
      TabIndex        =   2
      Top             =   9240
      Width           =   14535
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Exit"
         Height          =   375
         Left            =   13200
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cmbFld 
         Height          =   315
         ItemData        =   "frmItemView.frx":312D
         Left            =   120
         List            =   "frmItemView.frx":3140
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtVar 
         Height          =   285
         Left            =   2280
         TabIndex        =   4
         Top             =   240
         Width           =   2895
      End
      Begin VB.CommandButton cmdSearch 
         Height          =   495
         Left            =   5280
         Picture         =   "frmItemView.frx":3179
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   160
         Width           =   735
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFItems 
      Height          =   8295
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   14631
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   10080
      Width           =   14910
      _ExtentX        =   26300
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   26458
            MinWidth        =   26458
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item View"
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
      TabIndex        =   7
      Top             =   480
      Width           =   5295
   End
End
Attribute VB_Name = "frmItemView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Close All
    Unload Me
End Sub

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

Private Sub cmdSearch_Click()
On Error GoTo er_Eh:
    STRSQL = "Select * from InxV_Items Where " & Trim(cmbFld.Text) & " like '%" & Trim(txtVar.Text) & "%' Order By Item_ID"
    FMT_Grid
    Exit Sub
er_Eh:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub Form_Load()
    STRSQL = "Select * from InxV_Items Order By Item_ID"
    FMT_Grid
    cmbFld.Text = "SERIAL_NO"
End Sub

Private Sub MSFItems_DblClick()
    frmItemTrans.lblItemID.Caption = MSFItems.TextMatrix(MSFItems.Row, 0)
    Load frmItemTrans
    frmItemTrans.Show (1)
End Sub
