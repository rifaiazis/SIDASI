VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Begin VB.Form konversi_jarak 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Konversi Jarak"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   9660
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
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
      Height          =   615
      Left            =   4560
      TabIndex        =   65
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Lanjut"
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
      Left            =   2280
      TabIndex        =   64
      Top             =   4800
      Width           =   2055
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   11400
      OleObjectBlob   =   "konversi_jarak.frx":0000
      Top             =   360
   End
   Begin VB.Frame Frame1 
      Height          =   8775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9375
      Begin VB.CommandButton Command1 
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
         Height          =   615
         Left            =   7080
         TabIndex        =   62
         Top             =   7920
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3120
         TabIndex        =   61
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3120
         TabIndex        =   60
         Top             =   2040
         Width           =   2415
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   375
         Left            =   6000
         OleObjectBlob   =   "konversi_jarak.frx":0234
         TabIndex        =   58
         Top             =   2040
         Width           =   615
      End
      Begin VB.Frame FrameMM 
         Caption         =   " Satuan Hasil (Ke) "
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   6000
         TabIndex        =   51
         Top             =   2640
         Width           =   3015
         Begin VB.OptionButton Option42 
            Caption         =   "CM (Centimeter)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   57
            Top             =   2880
            Width           =   2175
         End
         Begin VB.OptionButton Option41 
            Caption         =   "DM (Desimeter)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   56
            Top             =   2400
            Width           =   2175
         End
         Begin VB.OptionButton Option40 
            Caption         =   "M (Meter)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   55
            Top             =   1920
            Width           =   2295
         End
         Begin VB.OptionButton Option39 
            Caption         =   "DAM (Dekameter)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   54
            Top             =   1440
            Width           =   2055
         End
         Begin VB.OptionButton Option38 
            Caption         =   "HM (Hektometer)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   53
            Top             =   960
            Width           =   2175
         End
         Begin VB.OptionButton Option37 
            Caption         =   "KM (Kilometer)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   52
            Top             =   480
            Width           =   2295
         End
      End
      Begin VB.Frame FrameCM 
         Caption         =   " Satuan Hasil (Ke) "
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   6000
         TabIndex        =   44
         Top             =   2640
         Width           =   3015
         Begin VB.OptionButton Option36 
            Caption         =   "MM (Milimeter)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   50
            Top             =   2880
            Width           =   2055
         End
         Begin VB.OptionButton Option35 
            Caption         =   "DM (Desimeter)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   49
            Top             =   2400
            Width           =   2175
         End
         Begin VB.OptionButton Option34 
            Caption         =   "M (Meter)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   48
            Top             =   1920
            Width           =   2055
         End
         Begin VB.OptionButton Option33 
            Caption         =   "DAM (Dekameter)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   47
            Top             =   1440
            Width           =   2175
         End
         Begin VB.OptionButton Option32 
            Caption         =   "HM (Hektometer)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   46
            Top             =   960
            Width           =   2175
         End
         Begin VB.OptionButton Option31 
            Caption         =   "KM (Kilometer)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   45
            Top             =   480
            Width           =   2175
         End
      End
      Begin VB.Frame FrameDM 
         Caption         =   " Satuan Hasil (Ke) "
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   6000
         TabIndex        =   37
         Top             =   2640
         Width           =   3015
         Begin VB.OptionButton Option30 
            Caption         =   "MM (Milimeter)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   43
            Top             =   2880
            Width           =   2055
         End
         Begin VB.OptionButton Option29 
            Caption         =   "CM (Centimeter)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   42
            Top             =   2400
            Width           =   2055
         End
         Begin VB.OptionButton Option28 
            Caption         =   "M (Meter)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   41
            Top             =   1920
            Width           =   1935
         End
         Begin VB.OptionButton Option27 
            Caption         =   "DAM (Dekameter)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   40
            Top             =   1440
            Width           =   2055
         End
         Begin VB.OptionButton Option26 
            Caption         =   "HM (Hektometer)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   39
            Top             =   960
            Width           =   2055
         End
         Begin VB.OptionButton Option25 
            Caption         =   "KM (Kilometer)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   38
            Top             =   480
            Width           =   2175
         End
      End
      Begin VB.Frame FrameM 
         Caption         =   " Satuan Hasil (Ke) "
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   6000
         TabIndex        =   30
         Top             =   2640
         Width           =   3015
         Begin VB.OptionButton Option24 
            Caption         =   "MM (Milimeter)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   36
            Top             =   2880
            Width           =   1935
         End
         Begin VB.OptionButton Option23 
            Caption         =   "CM (Centimeter)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   35
            Top             =   2400
            Width           =   1935
         End
         Begin VB.OptionButton Option22 
            Caption         =   "DM (Desimeter)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   34
            Top             =   1920
            Width           =   1935
         End
         Begin VB.OptionButton Option21 
            Caption         =   "DAM (Dekameter)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   33
            Top             =   1440
            Width           =   2055
         End
         Begin VB.OptionButton Option20 
            Caption         =   "HM (Hektometer)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   32
            Top             =   960
            Width           =   2055
         End
         Begin VB.OptionButton Option19 
            Caption         =   "KM (Kilometer)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   31
            Top             =   480
            Width           =   1935
         End
      End
      Begin VB.Frame FrameDAM 
         Caption         =   " Satuan Hasil (Ke) "
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   6000
         TabIndex        =   23
         Top             =   2640
         Width           =   3015
         Begin VB.OptionButton Option18 
            Caption         =   "MM (Milimeter)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   29
            Top             =   2880
            Width           =   2175
         End
         Begin VB.OptionButton Option17 
            Caption         =   "CM (Centimeter)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   28
            Top             =   2400
            Width           =   2175
         End
         Begin VB.OptionButton Option16 
            Caption         =   "DM (Desimeter)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   27
            Top             =   1920
            Width           =   2295
         End
         Begin VB.OptionButton Option15 
            Caption         =   "M (Meter)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   26
            Top             =   1440
            Width           =   2055
         End
         Begin VB.OptionButton Option14 
            Caption         =   "HM (Hektometer)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   25
            Top             =   960
            Width           =   2175
         End
         Begin VB.OptionButton Option13 
            Caption         =   "KM (Kilometer)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   24
            Top             =   480
            Width           =   2295
         End
      End
      Begin VB.Frame FrameHM 
         Caption         =   " Satuan Hasil (Ke) "
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   6000
         TabIndex        =   16
         Top             =   2640
         Width           =   3015
         Begin VB.OptionButton Option7 
            Caption         =   "KM (Kilometer)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   22
            Top             =   480
            Width           =   2175
         End
         Begin VB.OptionButton Option8 
            Caption         =   "DAM (Dekameter)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   21
            Top             =   960
            Width           =   2175
         End
         Begin VB.OptionButton Option9 
            Caption         =   "M (Meter)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   20
            Top             =   1440
            Width           =   2055
         End
         Begin VB.OptionButton Option10 
            Caption         =   "DM (Desimeter)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   19
            Top             =   1920
            Width           =   2175
         End
         Begin VB.OptionButton Option11 
            Caption         =   "CM (Centimeter)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   18
            Top             =   2400
            Width           =   2175
         End
         Begin VB.OptionButton Option12 
            Caption         =   "MM (Milimeter)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   17
            Top             =   2880
            Width           =   2175
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Tambahan"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   360
         TabIndex        =   14
         Top             =   7200
         Width           =   5175
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   615
            Left            =   240
            OleObjectBlob   =   "konversi_jarak.frx":0294
            TabIndex        =   15
            Top             =   480
            Width           =   4815
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Keterangan"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   360
         TabIndex        =   12
         Top             =   3960
         Width           =   5175
         Begin VB.PictureBox Picture1 
            Height          =   2295
            Left            =   360
            Picture         =   "konversi_jarak.frx":03A4
            ScaleHeight     =   2235
            ScaleWidth      =   4395
            TabIndex        =   13
            Top             =   480
            Width           =   4455
         End
      End
      Begin VB.Frame FrameKM 
         Caption         =   " Satuan Hasil (Ke) "
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   6000
         TabIndex        =   5
         Top             =   2640
         Width           =   3015
         Begin VB.OptionButton Option6 
            Caption         =   "MM (Milimeter)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   11
            Top             =   2880
            Width           =   2055
         End
         Begin VB.OptionButton Option5 
            Caption         =   "CM (Centimeter)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   10
            Top             =   2400
            Width           =   2055
         End
         Begin VB.OptionButton Option4 
            Caption         =   "DM (Desimeter)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   9
            Top             =   1920
            Width           =   2055
         End
         Begin VB.OptionButton Option3 
            Caption         =   "M (Meter)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   8
            Top             =   1440
            Width           =   2055
         End
         Begin VB.OptionButton Option2 
            Caption         =   "DAM (Dekameter)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   7
            Top             =   960
            Width           =   2055
         End
         Begin VB.OptionButton Option1 
            Caption         =   "HM (Hektometer)"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   6
            Top             =   480
            Width           =   2055
         End
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   6720
         TabIndex        =   4
         Text            =   "---- Satuan Awal ----"
         Top             =   2040
         Width           =   2295
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   1335
         Left            =   240
         OleObjectBlob   =   "konversi_jarak.frx":D474
         TabIndex        =   3
         Top             =   240
         Width           =   8895
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   495
         Left            =   360
         OleObjectBlob   =   "konversi_jarak.frx":D734
         TabIndex        =   2
         Top             =   2880
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   375
         Left            =   360
         OleObjectBlob   =   "konversi_jarak.frx":D7B4
         TabIndex        =   1
         Top             =   2160
         Width           =   2055
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   855
      Left            =   1560
      OleObjectBlob   =   "konversi_jarak.frx":D832
      TabIndex        =   59
      Top             =   2640
      Width           =   6015
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   855
      Left            =   840
      OleObjectBlob   =   "konversi_jarak.frx":D8A8
      TabIndex        =   63
      Top             =   3720
      Width           =   7335
   End
End
Attribute VB_Name = "konversi_jarak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Text1_Text2 As Single

Private Sub Combo1_Click()
Select Case Combo1.Text
    Case "KM (Kilometer)"
        FrameKM.Visible = True
        FrameHM.Visible = False
        FrameDAM.Visible = False
        FrameM.Visible = False
        FrameDM.Visible = False
        FrameCM.Visible = False
        FrameMM.Visible = False
        Text2.Text = " "
    Case "HM (Hektometer)"
        FrameKM.Visible = False
        FrameHM.Visible = True
        FrameDAM.Visible = False
        FrameM.Visible = False
        FrameDM.Visible = False
        FrameCM.Visible = False
        FrameMM.Visible = False
        Text2.Text = " "
    Case "DAM (Dekameter)"
        FrameKM.Visible = False
        FrameHM.Visible = False
        FrameDAM.Visible = True
        FrameM.Visible = False
        FrameDM.Visible = False
        FrameCM.Visible = False
        FrameMM.Visible = False
        Text2.Text = " "
    Case "M (Meter)"
        FrameKM.Visible = False
        FrameHM.Visible = False
        FrameDAM.Visible = False
        FrameM.Visible = True
        FrameDM.Visible = False
        FrameCM.Visible = False
        FrameMM.Visible = False
        Text2.Text = " "
    Case "DM (Desimeter)"
        FrameKM.Visible = False
        FrameHM.Visible = False
        FrameDAM.Visible = False
        FrameM.Visible = False
        FrameDM.Visible = True
        FrameCM.Visible = False
        FrameMM.Visible = False
        Text2.Text = " "
    Case "CM (Centimeter)"
        FrameKM.Visible = False
        FrameHM.Visible = False
        FrameDAM.Visible = False
        FrameM.Visible = False
        FrameDM.Visible = False
        FrameCM.Visible = True
        FrameMM.Visible = False
        Text2.Text = " "
    Case "MM (Milimeter)"
        FrameKM.Visible = False
        FrameHM.Visible = False
        FrameDAM.Visible = False
        FrameM.Visible = False
        FrameDM.Visible = False
        FrameCM.Visible = False
        FrameMM.Visible = True
        Text2.Text = " "
End Select
End Sub

Private Sub Command1_Click()
    xxx = MsgBox("Apakah anda yakin ingin keluar ?", vbOKCancel, "Informasi")
        If xxx = vbOK Then
        Unload Me
        menu.Enabled = True
        menu.Show
        End If
End Sub

Private Sub Command2_Click()
SkinLabel6.Visible = False
SkinLabel7.Visible = False
Command2.Visible = False
Command3.Visible = False
Text2.Enabled = False
Frame1.Visible = True
FrameKM.Visible = False
FrameHM.Visible = False
FrameDAM.Visible = False
FrameM.Visible = False
FrameDM.Visible = False
FrameCM.Visible = False
FrameMM.Visible = False
End Sub

Private Sub Command3_Click()
xxx = MsgBox("Apakah anda yakin ingin keluar ?", vbOKCancel, "Informasi")
        If xxx = vbOK Then
        Unload Me
        menu.Enabled = True
        menu.Show
        End If
End Sub

Private Sub Option1_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option1.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text * 10
    End If
End Sub

Private Sub Option11_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option11.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text * 10000
       End If
End Sub

Private Sub Option12_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option12.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text * 100000
       End If
End Sub

Private Sub Option13_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option13.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text / 100
       End If
End Sub

Private Sub Option14_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option14.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text / 10
       End If
End Sub

Private Sub Option15_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option15.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text * 10
       End If
End Sub

Private Sub Option16_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option16.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text * 100
       End If
End Sub

Private Sub Option17_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option17.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text * 1000
       End If
End Sub

Private Sub Option18_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option18.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text * 10000
       End If
End Sub

Private Sub Option19_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option19.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text / 1000
       End If
End Sub

Private Sub Option2_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option2.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text * 100
       End If
End Sub

Private Sub Option20_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option20.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text / 100
       End If
End Sub

Private Sub Option21_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option21.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text / 10
       End If
End Sub

Private Sub Option22_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option22.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text * 10
       End If
End Sub

Private Sub Option23_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option23.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text * 100
       End If
End Sub

Private Sub Option24_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option24.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text * 1000
       End If
End Sub

Private Sub Option25_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option25.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text / 10000
       End If
End Sub

Private Sub Option26_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option26.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text / 1000
       End If
End Sub

Private Sub Option27_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option27.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text / 100
       End If
End Sub

Private Sub Option28_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option28.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text / 10
       End If
End Sub

Private Sub Option29_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option29.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text * 10
       End If
End Sub

Private Sub Option3_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option3.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text * 1000
       End If
End Sub

Private Sub Option30_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option30.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text * 100
       End If
End Sub

Private Sub Option31_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option31.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text / 100000
       End If
End Sub

Private Sub Option32_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option32.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text / 10000
       End If
End Sub

Private Sub Option33_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option33.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text / 1000
       End If
End Sub

Private Sub Option34_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option34.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text / 100
       End If
End Sub

Private Sub Option35_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option35.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text / 10
       End If
End Sub

Private Sub Option36_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option36.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text * 10
       End If
End Sub

Private Sub Option37_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option37.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text / 1000000
       End If
End Sub

Private Sub Option38_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option38.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text / 100000
       End If
End Sub

Private Sub Option39_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option39.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text / 10000
       End If
End Sub

Private Sub Option4_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option4.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text * 10000
       End If
End Sub

Private Sub Option40_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option40.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text / 1000
       End If
End Sub

Private Sub Option41_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option41.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text / 100
       End If
End Sub

Private Sub Option42_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option42.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text / 10
       End If
End Sub

Private Sub Option5_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option5.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text * 100000
       End If
End Sub

Private Sub Option6_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option6.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text * 1000000
       End If
End Sub

Private Sub Option7_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option7.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text / 10
       End If
End Sub

Private Sub Option8_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option8.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text * 10
       End If
End Sub

Private Sub Option9_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option9.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text * 100
       End If
End Sub

Private Sub Option10_Click()
If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
    MsgBox " Silahkan masukkan angka  ", vbOKOnly, "Informasi"
    Text1.Text = ""
    Text1.SetFocus
    Option10.Value = False
    Else
       Text2.Text = Text1.Text
       Hasil = Text2.Text
       Text2.Text = Text1.Text * 1000
    End If
End Sub


Private Sub Form_Load()
menu.Enabled = False
menu.Visible = False
Frame1.Visible = False
FrameKM.Visible = False
FrameHM.Visible = False
FrameDAM.Visible = False
FrameM.Visible = False
FrameDM.Visible = False
FrameCM.Visible = False
FrameMM.Visible = False
Combo1.AddItem "KM (Kilometer)"
Combo1.AddItem "HM (Hektometer)"
Combo1.AddItem "DAM (Dekameter)"
Combo1.AddItem "M (Meter)"
Combo1.AddItem "DM (Desimeter)"
Combo1.AddItem "CM (Centimeter)"
Combo1.AddItem "MM (Milimeter)"
Skin1.LoadSkin App.Path & "\Dogmas2.skn"
Skin1.ApplySkin Me.hWnd
End Sub
