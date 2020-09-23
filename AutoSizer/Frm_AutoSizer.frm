VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Frm_AutoSizer 
   BackColor       =   &H00C0C0C0&
   Caption         =   "AutoSizer : Auto Resize Form Controls"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6855
   Icon            =   "Frm_AutoSizer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdForm2 
      Caption         =   "Photo Zoom"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Tag             =   "AutoSizer:y"
      Top             =   3840
      Width           =   1455
   End
   Begin MyAutoSizer.AutoSizer AutoSizer1 
      Left            =   3000
      Top             =   240
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2535
      Left            =   2640
      TabIndex        =   7
      Tag             =   "AutoSizer:HW"
      ToolTipText     =   "Masika .S. Elvas +254 724 688 172 maselv_e@yahoo.co.uk"
      Top             =   1440
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4471
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "Frm_AutoSizer.frx":57E2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "List3"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "Frm_AutoSizer.frx":57FE
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Image3(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Image3(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Image3(2)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Image3(3)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Image3(4)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "Frm_AutoSizer.frx":581A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Image3(5)"
      Tab(2).ControlCount=   1
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         Height          =   1590
         Left            =   -74880
         TabIndex        =   8
         Tag             =   "AutoSizer:HW"
         Top             =   600
         Width           =   3015
      End
      Begin VB.Image Image3 
         Height          =   405
         Index           =   5
         Left            =   -73560
         Picture         =   "Frm_AutoSizer.frx":5836
         Stretch         =   -1  'True
         Tag             =   "AutoSizer:C"
         ToolTipText     =   "Masika .S. Elvas +254 724 688 172 maselv_e@yahoo.co.uk"
         Top             =   1080
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   405
         Index           =   4
         Left            =   1680
         Picture         =   "Frm_AutoSizer.frx":6250
         Stretch         =   -1  'True
         Tag             =   "AutoSizer:C"
         ToolTipText     =   "Masika .S. Elvas +254 724 688 172 maselv_e@yahoo.co.uk"
         Top             =   960
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   525
         Index           =   3
         Left            =   120
         Picture         =   "Frm_AutoSizer.frx":6C6A
         Stretch         =   -1  'True
         Tag             =   "AutoSizer:Y"
         ToolTipText     =   "Masika .S. Elvas +254 724 688 172 maselv_e@yahoo.co.uk"
         Top             =   1680
         Width           =   600
      End
      Begin VB.Image Image3 
         Height          =   525
         Index           =   2
         Left            =   3120
         Picture         =   "Frm_AutoSizer.frx":7684
         Stretch         =   -1  'True
         Tag             =   "AutoSizer:X"
         ToolTipText     =   "Masika .S. Elvas +254 724 688 172 maselv_e@yahoo.co.uk"
         Top             =   360
         Width           =   600
      End
      Begin VB.Image Image3 
         Height          =   525
         Index           =   1
         Left            =   120
         Picture         =   "Frm_AutoSizer.frx":809E
         Stretch         =   -1  'True
         ToolTipText     =   "Masika .S. Elvas +254 724 688 172 maselv_e@yahoo.co.uk"
         Top             =   360
         Width           =   600
      End
      Begin VB.Image Image3 
         Height          =   525
         Index           =   0
         Left            =   3120
         Picture         =   "Frm_AutoSizer.frx":8AB8
         Stretch         =   -1  'True
         Tag             =   "AutoSizer:XY"
         ToolTipText     =   "Masika .S. Elvas +254 724 688 172 maselv_e@yahoo.co.uk"
         Top             =   1680
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Tag             =   "AutoSizer:xy"
      Top             =   4500
      Width           =   1215
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2535
      Left            =   6480
      TabIndex        =   3
      Tag             =   "AutoSizer:xh"
      Top             =   1440
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Tag             =   "AutoSizer:wy"
      Top             =   3960
      Width           =   4095
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2640
      TabIndex        =   1
      Tag             =   "AutoSizer:w"
      Text            =   "Combo1"
      ToolTipText     =   "Masika .S. Elvas +254 724 688 172 maselv_e@yahoo.co.uk"
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Tag             =   "AutoSizer:H"
      ToolTipText     =   "Masika .S. Elvas +254 724 688 172 maselv_e@yahoo.co.uk"
      Top             =   960
      Width           =   2415
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   2370
         Left            =   120
         TabIndex        =   4
         Tag             =   "AutoSizer:H"
         ToolTipText     =   "Masika .S. Elvas +254 724 688 172 maselv_e@yahoo.co.uk"
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   120
      Picture         =   "Frm_AutoSizer.frx":94D2
      Stretch         =   -1  'True
      ToolTipText     =   "Masika .S. Elvas +254 724 688 172 maselv_e@yahoo.co.uk"
      Top             =   120
      Width           =   600
   End
   Begin VB.Label lblDeveloper 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Masika .S. Elvas +254 724 688 172 maselv_e@yahoo.co.uk"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Tag             =   "AutoSizer:wy"
      Top             =   4530
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   645
      Left            =   6120
      Picture         =   "Frm_AutoSizer.frx":9EEC
      Stretch         =   -1  'True
      Tag             =   "AutoSizer:x"
      ToolTipText     =   "Masika .S. Elvas +254 724 688 172 maselv_e@yahoo.co.uk"
      Top             =   105
      Width           =   600
   End
   Begin VB.Image ImgFooter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   0
      Picture         =   "Frm_AutoSizer.frx":D5C7
      Stretch         =   -1  'True
      Tag             =   "AutoSizer:yw"
      ToolTipText     =   "Masika .S. Elvas +254 724 688 172 maselv_e@yahoo.co.uk"
      Top             =   4320
      Width           =   6855
   End
   Begin VB.Image ImgHeader 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   0
      Picture         =   "Frm_AutoSizer.frx":E07F
      Stretch         =   -1  'True
      Tag             =   "AutoSizer:w"
      ToolTipText     =   "Masika .S. Elvas +254 724 688 172 maselv_e@yahoo.co.uk"
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "Frm_AutoSizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdForm2_Click()
    Frm_AutoSizer2.Show vbModal, Me
End Sub
