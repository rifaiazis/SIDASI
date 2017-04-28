VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.MDIForm menu 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000B&
   Caption         =   "Aplikasi Ayo Belajar"
   ClientHeight    =   8220
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   14355
   LinkTopic       =   "MDIForm1"
   Picture         =   "Menu Utama.frx":0000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "Menu Utama.frx":3C56C
      Top             =   120
   End
   Begin VB.Menu menu 
      Caption         =   "Menu"
      Begin VB.Menu jarak 
         Caption         =   "Perhitungan Jarak"
      End
      Begin VB.Menu bangundatar 
         Caption         =   "Perhitungan Bangun Datar"
      End
      Begin VB.Menu ha 
         Caption         =   "-"
      End
      Begin VB.Menu keluar 
         Caption         =   "Keluar"
      End
   End
   Begin VB.Menu saya 
      Caption         =   "Tentang"
   End
End
Attribute VB_Name = "menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bangundatar_Click()
bangun_datar.Show
End Sub

Private Sub jarak_Click()
konversi_jarak.Show
End Sub

Private Sub keluar_Click()
 xxx = MsgBox("Apakah anda yakin ingin keluar ?", vbOKCancel, "Informasi")
        If xxx = vbOK Then
        End
        End If
End Sub

Private Sub MDIForm_Load()
Skin1.LoadSkin App.Path & "\Dogmas2.skn"
Skin1.ApplySkin Me.hWnd
End Sub

Private Sub saya_Click()
tentang_saya.Show
End Sub
