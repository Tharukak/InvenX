VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIHome 
   BackColor       =   &H8000000C&
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   4560
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIHome.frx":0000
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   60
      ImageHeight     =   60
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIHome.frx":27A5B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIHome.frx":284F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIHome.frx":28F11
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIHome.frx":299FF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   635
      ButtonWidth     =   3810
      ButtonHeight    =   1746
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Item Master"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Gatepass"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&GRN"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Reports"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   2640
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   10
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   35278
            MinWidth        =   35278
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu Mnu_MM1 
      Caption         =   "&File"
      Begin VB.Menu Mnu_MM1SM1 
         Caption         =   "&Item Creation"
      End
      Begin VB.Menu Mnu_MM1SM2 
         Caption         =   "Item Transaction"
      End
      Begin VB.Menu Mnu_MM1BK1 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_MM1SM4 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu Mnu_MM2 
      Caption         =   "&Configuration"
      Begin VB.Menu Mnui_MM2SM1 
         Caption         =   "Item &Category"
      End
      Begin VB.Menu Mnui_MM2SM2 
         Caption         =   "Item &Sub Category"
      End
      Begin VB.Menu Mnui_MM2SM3 
         Caption         =   "Item &Brand"
      End
      Begin VB.Menu Mnui_MM2SM4 
         Caption         =   "Item &Model"
      End
   End
   Begin VB.Menu Mnu_MM3 
      Caption         =   "&Settings"
      Begin VB.Menu Mnui_MM3SM1 
         Caption         =   "&Company Creation"
      End
      Begin VB.Menu Mnui_MM3SM2 
         Caption         =   "&SBU Creation"
      End
      Begin VB.Menu Mnui_MM3SM3 
         Caption         =   "&Plant Creation"
      End
      Begin VB.Menu Mnui_MM3SM4 
         Caption         =   "&Division Creation"
      End
   End
   Begin VB.Menu Mnu_MM4 
      Caption         =   "&Security Settings"
      Begin VB.Menu Mnu_MM4SM1 
         Caption         =   "Access Controle"
      End
      Begin VB.Menu Mnu_MM4SM2 
         Caption         =   "Category Controle"
      End
   End
End
Attribute VB_Name = "MDIHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    MDIHome.Caption = "Welcome to InvenX -Version Revolution - " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Mnu_MM1SM1_Click()
    MenuAccess 5
    If RightsMode = 1 Then
        Load frmItem
        frmItem.Show (1)
    End If
End Sub

Private Sub Mnu_MM1SM2_Click()
    MenuAccess 6
    If RightsMode = 1 Then
        Load frmItemView
        frmItemView.Show (1)
    End If
End Sub

Private Sub Mnu_MM1SM4_Click()
    Close All
    Unload Me
End Sub

Private Sub Mnui_MM2SM1_Click()
    MenuAccess 1
    If RightsMode = 1 Then
        Load frmIcat
        frmIcat.Show (1)
    End If
End Sub

Private Sub Mnui_MM2SM2_Click()
    MenuAccess 2
    If RightsMode = 1 Then
        Load frmISubCat
        frmISubCat.Show (1)
    End If
End Sub

Private Sub Mnui_MM2SM3_Click()
    MenuAccess 3
    If RightsMode = 1 Then
        Load frmBrand
        frmBrand.Show (1)
    End If
End Sub

Private Sub Mnui_MM2SM4_Click()
    MenuAccess 4
    If RightsMode = 1 Then
        Load frmModel
        frmModel.Show (1)
    End If
End Sub

Private Sub Mnui_MM3SM1_Click()
    MenuAccess 9
    If RightsMode = 1 Then
        Load frmCompany
        frmCompany.Show (1)
    End If
End Sub

Private Sub Mnui_MM3SM2_Click()
    MenuAccess 10
    If RightsMode = 1 Then
        Load frmSBU
        frmSBU.Show (1)
    End If
End Sub

Private Sub Mnui_MM3SM3_Click()
    MenuAccess 11
    If RightsMode = 1 Then
        Load frmPlant
        frmPlant.Show (1)
    End If
End Sub

Private Sub Mnui_MM3SM4_Click()
    MenuAccess 12
    If RightsMode = 1 Then
        Load frmDivision
        frmDivision.Show (1)
    End If
End Sub
