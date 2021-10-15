VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmGP 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gate Pass"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12690
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
   Picture         =   "frmGP.frx":0000
   ScaleHeight     =   7740
   ScaleWidth      =   12690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSenderView 
      Caption         =   "*.*"
      Height          =   255
      Left            =   5280
      TabIndex        =   53
      Top             =   2040
      Width           =   375
   End
   Begin VB.Frame Frame3 
      Height          =   5835
      Left            =   11640
      TabIndex        =   51
      Top             =   1590
      Width           =   915
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Exit"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   5400
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Printing Options"
      Height          =   1335
      Left            =   120
      TabIndex        =   48
      Top             =   6000
      Width           =   3855
      Begin VB.OptionButton Opt2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Direct to Window"
         Height          =   195
         Left            =   240
         TabIndex        =   50
         Top             =   840
         Width           =   3375
      End
      Begin VB.OptionButton Opt1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Direct to Printer"
         Height          =   195
         Left            =   240
         TabIndex        =   49
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   4080
      TabIndex        =   45
      Top             =   6000
      Width           =   1575
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H8000000A&
         Caption         =   "&Add"
         Height          =   375
         Left            =   240
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H8000000A&
         Caption         =   "&Delete"
         Height          =   375
         Left            =   240
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   720
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3375
      Left            =   5880
      TabIndex        =   44
      Top             =   3960
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   5953
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmdItemView 
      BackColor       =   &H00C0C0C0&
      Height          =   525
      Left            =   4560
      Picture         =   "frmGP.frx":1BB3
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Press to View Inventory Items"
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton cmdUpload 
      BackColor       =   &H00C0C0C0&
      Height          =   525
      Left            =   5160
      Picture         =   "frmGP.frx":1FF5
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Press to Upload Data"
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1680
      TabIndex        =   41
      Top             =   4080
      Width           =   2775
   End
   Begin VB.TextBox txtItemRemarks 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      TabIndex        =   32
      Top             =   5640
      Width           =   3855
   End
   Begin MSComCtl2.DTPicker DTPRetDate 
      Height          =   375
      Left            =   10080
      TabIndex        =   31
      Top             =   840
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   54198275
      CurrentDate     =   44214
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Height          =   315
      Left            =   5520
      TabIndex        =   29
      Top             =   840
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DataCombo3"
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7800
      TabIndex        =   18
      Top             =   3540
      Width           =   3615
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   1680
      TabIndex        =   17
      Top             =   2040
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   "DataCombo1"
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1680
      TabIndex        =   15
      Top             =   1320
      Width           =   10815
   End
   Begin VB.TextBox txtSenderName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      TabIndex        =   4
      Top             =   3480
      Width           =   3975
   End
   Begin MSComCtl2.DTPicker DTPTransDate 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   54198275
      CurrentDate     =   44214
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7485
      Width           =   12690
      _ExtentX        =   22384
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "2/3/2021"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "10:48 AM"
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
            Object.Width           =   14111
            MinWidth        =   14111
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   315
      Left            =   7800
      TabIndex        =   28
      Top             =   2040
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   "DataCombo1"
   End
   Begin VB.Label Label26 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Barcode Number"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label27 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Item Brand Name"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label28 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Item Model Name"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label Label29 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Item Model No."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label Label30 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Item Remarks"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label lblModelNumber 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   35
      Top             =   5280
      Width           =   3855
   End
   Begin VB.Label lblModelName 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   34
      Top             =   4920
      Width           =   3855
   End
   Begin VB.Label lblBrandName 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   33
      Top             =   4560
      Width           =   3855
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Return Expected Date"
      Height          =   255
      Left            =   8160
      TabIndex        =   30
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Receiver  Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6000
      TabIndex        =   27
      Top             =   1740
      Width           =   5415
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Company Name"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6000
      TabIndex        =   26
      Top             =   2100
      Width           =   1695
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Address Line 01"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6000
      TabIndex        =   25
      Top             =   2460
      Width           =   1695
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Address Line 02"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6000
      TabIndex        =   24
      Top             =   2820
      Width           =   1695
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Address Line 03"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6000
      TabIndex        =   23
      Top             =   3180
      Width           =   1695
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Contact Person"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6000
      TabIndex        =   22
      Top             =   3540
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   2250
      Left            =   5880
      Top             =   1680
      Width           =   5655
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   21
      Top             =   2460
      Width           =   3615
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   20
      Top             =   2820
      Width           =   3615
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   19
      Top             =   3180
      Width           =   3615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Gate Pass Method"
      Height          =   255
      Left            =   4080
      TabIndex        =   16
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Header Remarks"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Sender Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1740
      Width           =   5415
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Company Name"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2100
      Width           =   1335
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Address Line 01"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2460
      Width           =   1335
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Address Line 02"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2820
      Width           =   1335
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Address Line 03"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   3180
      Width           =   1335
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Contact Person"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3540
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   2250
      Left            =   120
      Top             =   1680
      Width           =   5655
   End
   Begin VB.Label lblAdd1 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   2460
      Width           =   3975
   End
   Begin VB.Label lblAdd2 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   2820
      Width           =   3975
   End
   Begin VB.Label lblAdd3 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   3180
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Date"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Gate Pass Creation"
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
      TabIndex        =   1
      Top             =   480
      Width           =   5295
   End
End
Attribute VB_Name = "frmGP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Close All
    Unload Me
End Sub

Private Sub Form_Load()
    DTPTransDate.Value = Date
    DTPRetDate.Value = Date
    Opt1.Value = True
End Sub
