VERSION 5.00
Begin VB.Form Frm_AutoSizer2 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Photo Zoom"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2535
   Icon            =   "Frm_AutoSizer2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   2535
   StartUpPosition =   1  'CenterOwner
   Begin MyAutoSizer.AutoSizer AutoSizer1 
      Left            =   0
      Top             =   2760
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Tag             =   "AutoSizer:xy"
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Tag             =   "AutoSizer:WH"
      Top             =   480
      Width           =   2055
      Begin VB.Image Image4 
         Height          =   2205
         Left            =   120
         Picture         =   "Frm_AutoSizer2.frx":57E2
         Stretch         =   -1  'True
         Tag             =   "AutoSizer:C"
         ToolTipText     =   "Masika .S. Elvas +254 724 688 172 maselv_e@yahoo.co.uk"
         Top             =   240
         Width           =   1800
      End
   End
   Begin VB.Image ImgHeader 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   0
      Picture         =   "Frm_AutoSizer2.frx":8EBD
      Stretch         =   -1  'True
      Tag             =   "AutoSizer:w"
      Top             =   0
      Width           =   2535
   End
   Begin VB.Image ImgFooter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   0
      Picture         =   "Frm_AutoSizer2.frx":9975
      Stretch         =   -1  'True
      Tag             =   "AutoSizer:yw"
      Top             =   3240
      Width           =   2535
   End
End
Attribute VB_Name = "Frm_AutoSizer2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
