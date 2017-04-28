VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form bangun_datar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Perhitungan Bangun Datar"
   ClientHeight    =   9525
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8445
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9525
   ScaleWidth      =   8445
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   44
      Top             =   5640
      Width           =   2175
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   855
      Left            =   600
      OleObjectBlob   =   "bangun_datar.frx":0000
      TabIndex        =   43
      Top             =   4320
      Width           =   7335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   855
      Left            =   1320
      OleObjectBlob   =   "bangun_datar.frx":015C
      TabIndex        =   42
      Top             =   3120
      Width           =   6015
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4920
      TabIndex        =   41
      Text            =   "------------ Pilih Bangun Datar ----------"
      Top             =   480
      Width           =   3255
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   360
      OleObjectBlob   =   "bangun_datar.frx":01D2
      Top             =   8040
   End
   Begin VB.Frame Frame1 
      Height          =   9255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      Begin VB.CommandButton Command3 
         Caption         =   "Keluar"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5160
         TabIndex        =   8
         Top             =   8520
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Kosongkan"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3120
         TabIndex        =   7
         Top             =   8520
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Hitung"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1080
         TabIndex        =   6
         Top             =   8520
         Width           =   1935
      End
      Begin VB.Frame Frame5 
         Caption         =   "Hasil Perhitungan Bangun Datar"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   4200
         TabIndex        =   5
         Top             =   5520
         Width           =   3855
         Begin ACTIVESKINLibCtl.SkinLabel Label1 
            Height          =   375
            Left            =   240
            OleObjectBlob   =   "bangun_datar.frx":1FEA9
            TabIndex        =   18
            Top             =   2040
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel Label3 
            Height          =   375
            Left            =   240
            OleObjectBlob   =   "bangun_datar.frx":1FF17
            TabIndex        =   17
            Top             =   1440
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel Label2 
            Height          =   375
            Left            =   240
            OleObjectBlob   =   "bangun_datar.frx":1FF7F
            TabIndex        =   16
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox Text8 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1680
            TabIndex        =   15
            Top             =   1440
            Width           =   1935
         End
         Begin VB.TextBox Text7 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1680
            TabIndex        =   14
            Top             =   2040
            Width           =   1935
         End
         Begin VB.TextBox Text6 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1680
            TabIndex        =   13
            Top             =   840
            Width           =   1935
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Rumus Perhitungan Bangun Datar"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4455
         Left            =   4200
         TabIndex        =   4
         Top             =   840
         Width           =   3855
         Begin VB.PictureBox Picture88 
            Height          =   1455
            Left            =   120
            Picture         =   "bangun_datar.frx":1FFDF
            ScaleHeight     =   1395
            ScaleWidth      =   3555
            TabIndex        =   40
            Top             =   1680
            Visible         =   0   'False
            Width           =   3615
         End
         Begin VB.PictureBox Picture77 
            Height          =   1575
            Left            =   120
            Picture         =   "bangun_datar.frx":26776
            ScaleHeight     =   1515
            ScaleWidth      =   3555
            TabIndex        =   39
            Top             =   1680
            Visible         =   0   'False
            Width           =   3615
         End
         Begin VB.PictureBox Picture66 
            Height          =   1935
            Left            =   120
            Picture         =   "bangun_datar.frx":2C594
            ScaleHeight     =   1875
            ScaleWidth      =   3555
            TabIndex        =   38
            Top             =   1440
            Visible         =   0   'False
            Width           =   3615
         End
         Begin VB.PictureBox Picture55 
            Height          =   1935
            Left            =   120
            Picture         =   "bangun_datar.frx":3310E
            ScaleHeight     =   1875
            ScaleWidth      =   3555
            TabIndex        =   37
            Top             =   1440
            Visible         =   0   'False
            Width           =   3615
         End
         Begin VB.PictureBox Picture44 
            Height          =   1575
            Left            =   120
            Picture         =   "bangun_datar.frx":35FFC
            ScaleHeight     =   1515
            ScaleWidth      =   3555
            TabIndex        =   36
            Top             =   1560
            Visible         =   0   'False
            Width           =   3615
         End
         Begin VB.PictureBox Picture33 
            Height          =   2055
            Left            =   120
            Picture         =   "bangun_datar.frx":3C88B
            ScaleHeight     =   1995
            ScaleWidth      =   3555
            TabIndex        =   35
            Top             =   1320
            Visible         =   0   'False
            Width           =   3615
         End
         Begin VB.PictureBox Picture22 
            Height          =   1575
            Left            =   240
            Picture         =   "bangun_datar.frx":42E93
            ScaleHeight     =   1515
            ScaleWidth      =   3315
            TabIndex        =   34
            Top             =   1560
            Visible         =   0   'False
            Width           =   3375
         End
         Begin VB.PictureBox Picture11 
            Height          =   2055
            Left            =   360
            Picture         =   "bangun_datar.frx":48EA2
            ScaleHeight     =   1995
            ScaleWidth      =   3075
            TabIndex        =   33
            Top             =   1200
            Visible         =   0   'False
            Width           =   3135
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ukuran Bangun Datar"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   120
         TabIndex        =   3
         Top             =   5520
         Width           =   3855
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1680
            TabIndex        =   32
            Top             =   360
            Width           =   1935
         End
         Begin ACTIVESKINLibCtl.SkinLabel Label8 
            Height          =   375
            Left            =   120
            OleObjectBlob   =   "bangun_datar.frx":4BA8C
            TabIndex        =   23
            Top             =   2280
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel Label7 
            Height          =   375
            Left            =   120
            OleObjectBlob   =   "bangun_datar.frx":4BAE4
            TabIndex        =   22
            Top             =   1800
            Width           =   1455
         End
         Begin ACTIVESKINLibCtl.SkinLabel Label6 
            Height          =   375
            Left            =   120
            OleObjectBlob   =   "bangun_datar.frx":4BB3C
            TabIndex        =   21
            Top             =   1320
            Width           =   1455
         End
         Begin ACTIVESKINLibCtl.SkinLabel Label5 
            Height          =   375
            Left            =   120
            OleObjectBlob   =   "bangun_datar.frx":4BB94
            TabIndex        =   20
            Top             =   840
            Width           =   1455
         End
         Begin ACTIVESKINLibCtl.SkinLabel Label4 
            Height          =   375
            Left            =   120
            OleObjectBlob   =   "bangun_datar.frx":4BBEC
            TabIndex        =   19
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox Text5 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1680
            TabIndex        =   12
            Top             =   2280
            Width           =   1935
         End
         Begin VB.TextBox Text4 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1680
            TabIndex        =   11
            Top             =   1800
            Width           =   1935
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1680
            TabIndex        =   10
            Top             =   1320
            Width           =   1935
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1680
            TabIndex        =   9
            Top             =   840
            Width           =   1935
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Gambar Bangun Datar"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4455
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   3855
         Begin VB.PictureBox Picture8 
            Height          =   3255
            Left            =   120
            Picture         =   "bangun_datar.frx":4BC44
            ScaleHeight     =   3195
            ScaleWidth      =   3555
            TabIndex        =   31
            Top             =   720
            Visible         =   0   'False
            Width           =   3615
         End
         Begin VB.PictureBox Picture7 
            Height          =   2895
            Left            =   240
            Picture         =   "bangun_datar.frx":506CA
            ScaleHeight     =   2835
            ScaleWidth      =   3195
            TabIndex        =   30
            Top             =   840
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.PictureBox Picture6 
            Height          =   3735
            Left            =   720
            Picture         =   "bangun_datar.frx":548B7
            ScaleHeight     =   3675
            ScaleWidth      =   2355
            TabIndex        =   29
            Top             =   360
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.PictureBox Picture5 
            Height          =   3855
            Left            =   600
            Picture         =   "bangun_datar.frx":5F9DB
            ScaleHeight     =   3795
            ScaleWidth      =   2715
            TabIndex        =   28
            Top             =   360
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.PictureBox Picture4 
            Height          =   3375
            Left            =   480
            Picture         =   "bangun_datar.frx":64D2E
            ScaleHeight     =   3315
            ScaleWidth      =   2955
            TabIndex        =   27
            Top             =   600
            Visible         =   0   'False
            Width           =   3015
         End
         Begin VB.PictureBox Picture3 
            Height          =   3975
            Left            =   480
            Picture         =   "bangun_datar.frx":68DE3
            ScaleHeight     =   3915
            ScaleWidth      =   2835
            TabIndex        =   26
            Top             =   360
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.PictureBox Picture2 
            Height          =   3855
            Left            =   720
            Picture         =   "bangun_datar.frx":6D508
            ScaleHeight     =   3795
            ScaleWidth      =   2475
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   360
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.PictureBox Picture1 
            Height          =   3480
            Left            =   360
            Picture         =   "bangun_datar.frx":70D1D
            ScaleHeight     =   228
            ScaleMode       =   0  'User
            ScaleWidth      =   207
            TabIndex        =   24
            Top             =   480
            Visible         =   0   'False
            Width           =   3165
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel Label9 
         Height          =   495
         Left            =   120
         OleObjectBlob   =   "bangun_datar.frx":74005
         TabIndex        =   1
         Top             =   240
         Width           =   4455
      End
   End
End
Attribute VB_Name = "bangun_datar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim q, w, e, r, t, l, sm, k As Integer
Private Sub bersih()
Label4.Caption = ""
Label5.Caption = ""
Label6.Caption = ""
Label7.Caption = ""
Label8.Caption = ""
Label9.Caption = ""
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text1.Visible = False
Text2.Visible = False
Text3.Visible = False
Text4.Visible = False
Text5.Visible = False
Text6.Visible = False
Text7.Visible = False
Text8.Visible = False
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = False
Picture8.Visible = False
Picture11.Visible = False
Picture22.Visible = False
Picture33.Visible = False
Picture44.Visible = False
Picture55.Visible = False
Picture66.Visible = False
Picture77.Visible = False
Picture88.Visible = False
End Sub
Private Sub kelihatan1()
Label4.Visible = False
Label5.Visible = False
Label7.Visible = False
Label8.Visible = False
Text1.Visible = False
Text2.Visible = False
Text4.Visible = False
Text5.Visible = False
End Sub
Private Sub kelihatan2()
Label4.Visible = False
Label6.Visible = False
Label8.Visible = False
Text1.Visible = False
Text3.Visible = False
Text5.Visible = False


End Sub
Private Sub kelihatanall()
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Label7.Visible = True
Label8.Visible = True
Text1.Visible = True
Text2.Visible = True
Text3.Visible = True
Text4.Visible = True
Text5.Visible = True
End Sub
Private Sub kelihatan4()
Label6.Visible = False
Text3.Visible = False
End Sub

Private Sub kelihatan5()
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Label7.Visible = True
Label8.Visible = True
Text1.Visible = False
Text2.Visible = True
Text3.Visible = True
Text4.Visible = True
Text5.Visible = False
End Sub

Private Sub Combo1_Click()
Select Case Combo1.Text
Case "Persegi"
Call bersih
Call kelihatanall
Call kelihatan2
Label9.Caption = "Persegi"
Label5.Caption = " Sisi  "
Label7.Caption = " Sisi  "
SkinLabel1.Visible = False
SkinLabel2.Visible = False
Command4.Visible = False
Frame1.Visible = True
Label1.Visible = False
Text6.Visible = True
Text8.Visible = True
Text7.Visible = False
Label2.Visible = True
Label3.Visible = True
Picture1.Visible = True
Picture11.Visible = True
Case "Persegi Panjang"
Call bersih
Call kelihatanall
Call kelihatan2
Label9.Caption = "Persegi Panjang"
Label5.Caption = " Panjang  "
Label7.Caption = " Lebar  "
SkinLabel1.Visible = False
SkinLabel2.Visible = False
Command4.Visible = False
Frame1.Visible = True
Text6.Visible = True
Text8.Visible = True
Label2.Visible = True
Label3.Visible = True
Label1.Visible = False
Text7.Visible = False
Picture2.Visible = True
Picture22.Visible = True
Case "Segitiga"
Call bersih
Call kelihatanall
Call kelihatan2
Label9.Caption = "Segitiga"
Label5.Caption = " Alas  "
Label7.Caption = " Tinggi  "
SkinLabel1.Visible = False
SkinLabel2.Visible = False
Command4.Visible = False
Frame1.Visible = True
Text6.Visible = True
Text8.Visible = True
Label2.Visible = True
Label3.Visible = True
Label1.Visible = True
Text7.Visible = True
Picture3.Visible = True
Picture33.Visible = True
Case "Lingkarang"
Call bersih
Call kelihatanall
Call kelihatan1
Label9.Caption = "Lingkaran"
Label6.Caption = " Jari-jari  "
SkinLabel1.Visible = False
SkinLabel2.Visible = False
Command4.Visible = False
Frame1.Visible = True
Text6.Visible = True
Text8.Visible = True
Label2.Visible = True
Label3.Visible = True
Label1.Visible = False
Text7.Visible = False
Picture4.Visible = True
Picture44.Visible = True
Case "Layang-layang"
Call bersih
Call kelihatanall
Call kelihatan4
Label9.Caption = "Layang-layang"
Label4.Caption = " Sisi Panjang  "
Label5.Caption = " Sisi Pendek  "
Label7.Caption = " Diagonal 1  "
Label8.Caption = " Diagonal 2  "
SkinLabel1.Visible = False
SkinLabel2.Visible = False
Command4.Visible = False
Frame1.Visible = True
Text6.Visible = True
Text8.Visible = True
Label2.Visible = True
Label3.Visible = True
Label1.Visible = False
Text7.Visible = False
Picture6.Visible = True
Picture66.Visible = True
Case "Belah Ketupat"
Call bersih
Call kelihatanall
Call kelihatan2
Label9.Caption = "Belah Ketupat"
Label5.Caption = " Diagonal 1  "
Label7.Caption = " Diagonal 2  "
SkinLabel1.Visible = False
SkinLabel2.Visible = False
Command4.Visible = False
Frame1.Visible = True
Text6.Visible = True
Text8.Visible = True
Label2.Visible = True
Label3.Visible = True
Label1.Visible = True
Text7.Visible = True
Picture5.Visible = True
Picture55.Visible = True
Case "Jajar Genjang"
Call bersih
Call kelihatan5
Label9.Caption = "Jajar Genjang"
Label5.Caption = " Alas  "
Label6.Caption = " Tinggi  "
Label7.Caption = " Sisi Miring  "
SkinLabel1.Visible = False
SkinLabel2.Visible = False
Command4.Visible = False
Frame1.Visible = True
Text6.Visible = True
Text8.Visible = True
Label2.Visible = True
Label3.Visible = True
Label1.Visible = False
Text7.Visible = False
Picture7.Visible = True
Picture77.Visible = True
Case "Trapesium"
Call bersih
Call kelihatanall
Label9.Caption = "Trapesium"
Label4.Caption = " Sisi Panjang  "
Label5.Caption = " Sisi Pendek  "
Label6.Caption = " Tinggi  "
Label7.Caption = " Miring Kanan  "
Label8.Caption = " Miring Kiri  "
SkinLabel1.Visible = False
SkinLabel2.Visible = False
Command4.Visible = False
Frame1.Visible = True
Text6.Visible = True
Text8.Visible = True
Label2.Visible = True
Label3.Visible = True
Label1.Visible = False
Text7.Visible = False
Picture8.Visible = True
Picture88.Visible = True
End Select
End Sub


Private Sub Command1_Click()
q = Val(Text1)
w = Val(Text2)
e = Val(Text3)
r = Val(Text4)
t = Val(Text5)
Select Case Combo1.Text
Case "Persegi"
Call kelihatanall
Call kelihatan2
Label9.Caption = "Persegi"
Label5.Caption = " Sisi  "
Label7.Caption = " Sisi  "
If Text2.Text = "" Or Not IsNumeric(Text2.Text) Then
    MsgBox " Silahkan isi dengan angka ", vbOKOnly, "Informasi"
    Text2.Text = ""
    Text2.SetFocus
    Else
    If Text4.Text = "" Or Not IsNumeric(Text4.Text) Then
    MsgBox " Silahkan isi dengan angka ", vbOKOnly, "Informasi"
    Text4.Text = ""
    Text4.SetFocus
    Else
    If w > r Or w < r Then
    MsgBox " Kedua sisi persegi harus sama  ", vbOKOnly, "Informasi"
    Text4.Text = ""
    Text4.SetFocus
    Else
    l = w * r
    k = 4 * w
    Text6.Text = " " & l & " Cm2 "
    Text8.Text = " " & k & " Cm "
    End If
    End If
    End If
    
Case "Persegi Panjang"
Call kelihatanall
Call kelihatan2
Label9.Caption = "Persegi Panjang"
Label5.Caption = " Panjang  "
Label7.Caption = " Lebar  "
If Text2.Text = "" Or Not IsNumeric(Text2.Text) Then
    MsgBox " Silahkan isi dengan angka ", vbOKOnly, "Informasi"
    Text2.Text = ""
    Text2.SetFocus
    Else
    If Text4.Text = "" Or Not IsNumeric(Text4.Text) Then
    MsgBox " Silahkan isi dengan angka ", vbOKOnly, "Informasi"
    Text4.Text = ""
    Text4.SetFocus
    Else
    If w = r Then
    MsgBox " Kedua sisi harus berbeda ", vbOKOnly, "Informasi"
    Text4.Text = ""
    Text4.SetFocus
    Else
    l = w * r
    k = 2 * (w + r)
    Text6.Text = " " & l & " Cm2 "
    Text8.Text = " " & k & " Cm "
    End If
    End If
    End If
Case "Segitiga"
Call kelihatanall
Call kelihatan2
Label9.Caption = "Segitiga"
Label5.Caption = " Alas  "
Label7.Caption = " Tinggi  "
If Text2.Text = "" Or Not IsNumeric(Text2.Text) Then
    MsgBox " Silahkan isi dengan angka ", vbOKOnly, "Informasi"
    Text2.Text = ""
    Text2.SetFocus
    Else
    If Text4.Text = "" Or Not IsNumeric(Text4.Text) Then
    MsgBox " Silahkan isi dengan angka ", vbOKOnly, "Informasi"
    Text4.Text = ""
    Text4.SetFocus
    Else
    l = 0.5 * w * r
    sm = Math.Sqr((0.5 * w) ^ 2 + r ^ 2)
    k = w * (2 * sm)
    Text6.Text = " " & l & " Cm2 "
    Text7.Text = " " & sm & " Cm "
    Text8.Text = " " & k & " Cm "
    End If
    End If
Case "Lingkarang"
Call kelihatanall
Call kelihatan1
Label9.Caption = "Lingkaran"
Label6.Caption = " Jari-jari  "
If Text3 = "" Or Not IsNumeric(Text3.Text) Then
MsgBox "Masukanlah dengan angka", vbOKOnly
Text3.Text = ""
Text3.SetFocus
Else
If e Mod 7 = 0 Then
k = 2 * 22 / 7 * e
l = 22 / 7 * e ^ 2
Text6.Text = " " & l & " Cm2 "
Text8.Text = " " & k & " Cm "
    Else
    k = 2 * 3.14 * e
    l = 3.14 * e ^ 2
    Text6.Text = " " & l & " Cm2 "
    Text8.Text = " " & k & " Cm "
    End If
End If

Case "Layang-layang"
Call kelihatanall
Call kelihatan4
Label9.Caption = "Layang-layang"
Label4.Caption = " Sisi Panjang  "
Label5.Caption = " Sisi Pendek  "
Label7.Caption = " Diagonal 1  "
Label8.Caption = " Diagonal 2  "
If Text1 = "" Or Not IsNumeric(Text1.Text) Then
            MsgBox " Silahkan isi dengan angka  ", vbOKOnly, "Informasi"
            Text1.SetFocus
            Text1.Text = ""
        Else
        If Text2 = "" Or Not IsNumeric(Text2.Text) Then
            MsgBox " Silahkan isi dengan angka  ", vbOKOnly, "Informasi"
            Text2.SetFocus
            Text2.Text = ""
        Else
        If Text4 = "" Or Not IsNumeric(Text4.Text) Then
            MsgBox " Silahkan isi dengan angka  ", vbOKOnly, "Informasi"
            Text4.SetFocus
            Text4.Text = ""
        Else
        If Text5 = "" Or Not IsNumeric(Text5.Text) Then
            MsgBox " Silahkan isi dengan angka  ", vbOKOnly, "Informasi"
            Text5.SetFocus
            Text5.Text = ""
        Else
        If q <= w Then
        MsgBox "Sisi panjang harus lebih panjang dari sisi pendek  ", vbOKOnly, " Informasi  "
        Text2.SetFocus
        Text2 = ""
         Else
        If r <= t Then
        MsgBox "Diagonal 1 harus lebih panjang dari diagonal 2  ", vbOKOnly, " Informasi  "
        Text4.SetFocus
        Text4 = ""
        Else
            k = 2 * (q + w)
            l = 0.5 * (r * t)
            
            Text6.Text = " " & l & " Cm2 "
            Text8.Text = " " & k & " Cm "
        End If
        End If
        End If
        End If
        End If
        End If
Case "Belah Ketupat"
Call kelihatanall
Call kelihatan2
Label9.Caption = "Belah Ketupat"
Label5.Caption = " Diagonal 1  "
Label7.Caption = " Diagonal 2  "
If Text2.Text = "" Or Not IsNumeric(Text2.Text) Then
    MsgBox " Silahkan isi dengan angka ", vbOKOnly, "Informasi"
    Text2.Text = ""
    Text2.SetFocus
    Else
    If Text4.Text = "" Or Not IsNumeric(Text4.Text) Then
    MsgBox " Silahkan isi dengan angka ", vbOKOnly, "Informasi"
    Text4.Text = ""
    Text4.SetFocus
    Else
    If w > r Or w < r Then
    MsgBox " Kedua sisi persegi harus sama  ", vbOKOnly, "Informasi"
    Text4.Text = ""
    Text4.SetFocus
    Else
    l = 0.5 * w * r
    sm = Math.Sqr((0.5 * w) ^ 2 + (0.5 * r) ^ 2)
    k = 4 * sm
    Text6.Text = " " & l & " Cm2 "
    Text7.Text = " " & sm & " Cm "
    Text8.Text = " " & k & " Cm "
    End If
    End If
    End If
Case "Jajar Genjang"
Call kelihatan5
Label9.Caption = "Jajar Genjang"
Label5.Caption = " Alas  "
Label6.Caption = " Tinggi  "
Label7.Caption = " Sisi Miring  "
If Text2 = "" Or Not IsNumeric(Text2.Text) Then
    MsgBox " Silahkan isi dengan angka ", vbOKOnly, "Informasi "
    Text2.SetFocus
    Text2.Text = ""
    Else
    If Text3 = "" Or Not IsNumeric(Text3.Text) Then
    MsgBox " Silahkan isi dengan angka ", vbOKOnly, "Informasi "
    Text3.SetFocus
    Text3.Text = ""
    Else
    If Text4 = "" Or Not IsNumeric(Text4.Text) Then
    MsgBox " Silahkan isi dengan angka ", vbOKOnly, "Informasi "
    Text4.SetFocus
    Text4.Text = ""
    Else
    k = (2 * w) + (2 * r)
    l = w * e
    Text6.Text = " " & l & " Cm2 "
    Text8.Text = " " & k & " Cm "

End If
End If
End If

Case "Trapesium"
Call kelihatanall
Label9.Caption = "Trapesium"
Label4.Caption = " Sisi Panjang  "
Label5.Caption = " Sisi Pendek  "
Label6.Caption = " Tinggi  "
Label7.Caption = " Miring Kanan  "
Label8.Caption = " Miring Kiri  "
If Text1 = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan isi dengan angka ", vbOKOnly, "Informasi "
    Text1.SetFocus
    Text1.Text = ""
    Else
    If Text2 = "" Or Not IsNumeric(Text2.Text) Then
    MsgBox " Silahkan isi dengan angka ", vbOKOnly, "Informasi "
    Text2.SetFocus
    Text2.Text = ""
    Else
    If Text3 = "" Or Not IsNumeric(Text3.Text) Then
    MsgBox " Silahkan isi dengan angka ", vbOKOnly, "Informasi "
    Text3.SetFocus
    Text3.Text = ""
    Else
    If Text4 = "" Or Not IsNumeric(Text4.Text) Then
    MsgBox " Silahkan isi dengan angka ", vbOKOnly, "Informasi "
    Text4.SetFocus
    Text4.Text = ""
    Else
    If Text5 = "" Or Not IsNumeric(Text5.Text) Then
    MsgBox " Silahkan isi dengan angka ", vbOKOnly, "Informasi "
    Text5.SetFocus
    Text5.Text = ""
    Else
    If q < w Then
        MsgBox " Sisi atas harus lebih pendek dari sisi bawah  ", vbOKOnly, "Informasi"
        Text1 = ""
        Text2 = ""
        Text1.SetFocus
    Else
    If r = t Then
        MsgBox " Sisi kanan dan kiri tidak boleh sama  ", vbOKOnly, "Informasi"
        Text4 = ""
        Text5 = ""
        Text4.SetFocus
    Else
    If e > r Then
        MsgBox " Sisi kanan dan kiri harus lebih panjang dari tinggi trapesium  ", vbOKOnly, "Informasi"
        Text4 = ""
        Text5 = ""
        Text5.SetFocus
    Else
    k = q + w + r + t
    l = 0.5 * (q + w) * e
    Text6.Text = " " & l & " Cm2 "
    Text8.Text = " " & k & " Cm "
    End If
End If
End If
End If
End If
End If
End If
End If

End Select
End Sub

Private Sub Command2_Click()
Call bersih
End Sub

Private Sub Command3_Click()
 xxx = MsgBox("Apakah anda yakin ingin keluar ?", vbOKCancel, "Informasi")
        If xxx = vbOK Then
        Unload Me
        menu.Enabled = True
        menu.Show
        End If
End Sub

Private Sub Command4_Click()
 xxx = MsgBox("Apakah anda yakin ingin keluar ?", vbOKCancel, "Informasi")
        If xxx = vbOK Then
        Unload Me
        menu.Enabled = True
        menu.Show
        End If
End Sub

Private Sub Form_Load()
Frame1.Visible = False
Call bersih
Combo1.AddItem "Persegi"
Combo1.AddItem "Persegi Panjang"
Combo1.AddItem "Segitiga"
Combo1.AddItem "Lingkarang"
Combo1.AddItem "Layang-layang"
Combo1.AddItem "Belah Ketupat"
Combo1.AddItem "Jajar Genjang"
Combo1.AddItem "Trapesium"
Skin1.LoadSkin App.Path & "\Dogmas2.skn"
Skin1.ApplySkin Me.hWnd
menu.Enabled = False
menu.Visible = False
End Sub
