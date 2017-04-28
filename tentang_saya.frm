VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form tentang_saya 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tentang"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   5355
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   120
      Picture         =   "tentang_saya.frx":0000
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5280
      OleObjectBlob   =   "tentang_saya.frx":38FD
      Top             =   3360
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   4
      Top             =   3360
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   375
      Left            =   1200
      OleObjectBlob   =   "tentang_saya.frx":3B31
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   495
      Left            =   1200
      OleObjectBlob   =   "tentang_saya.frx":3B9F
      TabIndex        =   1
      Top             =   720
      Width           =   3975
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   1575
      Left            =   1200
      OleObjectBlob   =   "tentang_saya.frx":3C13
      TabIndex        =   2
      Top             =   1320
      Width           =   3975
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   615
      Left            =   120
      OleObjectBlob   =   "tentang_saya.frx":3E35
      TabIndex        =   3
      Top             =   3360
      Width           =   2895
   End
End
Attribute VB_Name = "tentang_saya"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
menu.Show
menu.Enabled = True
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & "\Dogmas2.skn"
Skin1.ApplySkin Me.hWnd
menu.Enabled = False
End Sub

