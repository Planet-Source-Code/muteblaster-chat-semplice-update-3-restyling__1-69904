VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form chat 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "chat"
   ClientHeight    =   10515
   ClientLeft      =   165
   ClientTop       =   840
   ClientWidth     =   15180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "chat.frx":0000
   ScaleHeight     =   10515
   ScaleWidth      =   15180
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer_chiusura_frmrooms 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   5160
      Top             =   360
   End
   Begin VB.Timer Timer_setparent_frmclient 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5520
      Top             =   360
   End
   Begin VB.Timer Timer_setparent_frmrooms 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5880
      Top             =   360
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H80000003&
      Caption         =   "altri canali"
      Height          =   4095
      Left            =   12840
      TabIndex        =   102
      Top             =   5280
      Width           =   2175
      Begin VB.PictureBox Picture23 
         Height          =   3735
         Left            =   120
         ScaleHeight     =   3675
         ScaleWidth      =   1875
         TabIndex        =   104
         Top             =   4800
         Width           =   1935
      End
      Begin VB.PictureBox Picture22 
         Height          =   3735
         Left            =   120
         ScaleHeight     =   3675
         ScaleWidth      =   1875
         TabIndex        =   103
         Top             =   4680
         Width           =   1935
      End
   End
   Begin client.CandyButton Cmddatabase_immagini 
      Height          =   375
      Left            =   5280
      TabIndex        =   101
      Top             =   1800
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "database immagini"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Timer Timer_frame6 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   6240
      Top             =   360
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H80000013&
      Height          =   1455
      Left            =   480
      TabIndex        =   96
      Top             =   10560
      Width           =   3255
      Begin client.CandyButton cmdchiudi_frame6 
         Height          =   255
         Left            =   2640
         TabIndex        =   99
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "x"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   16760976
         ColorButtonUp   =   15309136
         ColorButtonDown =   15309136
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   0
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H8000000A&
         Height          =   375
         Left            =   360
         TabIndex        =   98
         Text            =   "puoi spostare i bottoni dove vuoi"
         Top             =   600
         Width           =   2535
      End
      Begin client.chameleonButton chameleonButton3 
         Height          =   1335
         Left            =   0
         TabIndex        =   97
         Top             =   0
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   2355
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "chat.frx":24CEE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.CheckBox Check_muovi_componenti 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   95
      Top             =   10080
      Width           =   255
   End
   Begin client.CandyButton Cmdplayer 
      Height          =   375
      Left            =   6600
      TabIndex        =   94
      ToolTipText     =   "audio player"
      Top             =   8400
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "chat.frx":24D0A
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin client.CandyButton Cmd_animazioni_flash 
      Height          =   375
      Left            =   7440
      TabIndex        =   93
      ToolTipText     =   "animazioni flash"
      Top             =   8400
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "chat.frx":259E4
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin client.CandyButton Cmdtxtchat_color 
      Height          =   375
      Left            =   3480
      TabIndex        =   92
      ToolTipText     =   "cambia sfondo del testo ricevuto"
      Top             =   8400
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "chat.frx":266BE
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin client.CandyButton Cmdtxtsend_color 
      Height          =   375
      Left            =   2640
      TabIndex        =   91
      ToolTipText     =   "cambia sfondo del testo da inviare"
      Top             =   8400
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "chat.frx":27398
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin client.CandyButton Cmdunderline 
      Height          =   255
      Left            =   3000
      TabIndex        =   90
      ToolTipText     =   "underline"
      Top             =   8955
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "U"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin client.CandyButton Cmditalic 
      Height          =   255
      Left            =   2640
      TabIndex        =   89
      ToolTipText     =   "italic"
      Top             =   8955
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "I"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin client.CandyButton Cmdbold 
      Height          =   255
      Left            =   2280
      TabIndex        =   88
      ToolTipText     =   "bold"
      Top             =   8955
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "B"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.PictureBox Picture21 
      BackColor       =   &H80000013&
      Height          =   8055
      Left            =   8880
      ScaleHeight     =   7995
      ScaleWidth      =   5475
      TabIndex        =   87
      Top             =   10800
      Width           =   5535
   End
   Begin VB.PictureBox Picture20 
      BackColor       =   &H80000013&
      Height          =   6135
      Left            =   5880
      ScaleHeight     =   6075
      ScaleWidth      =   4395
      TabIndex        =   86
      Top             =   10800
      Width           =   4455
   End
   Begin VB.PictureBox Picture19 
      BackColor       =   &H80000013&
      Height          =   2295
      Left            =   3360
      ScaleHeight     =   2235
      ScaleWidth      =   4635
      TabIndex        =   85
      Top             =   10800
      Width           =   4695
   End
   Begin client.CandyButton Cmdcolor 
      Height          =   375
      Left            =   1920
      TabIndex        =   84
      ToolTipText     =   "messaggi colorati"
      Top             =   8400
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "chat.frx":28072
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.PictureBox Picture18 
      BackColor       =   &H80000013&
      Height          =   2775
      Left            =   10680
      ScaleHeight     =   2715
      ScaleWidth      =   4275
      TabIndex        =   82
      Top             =   10800
      Width           =   4335
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   1320
         TabIndex        =   83
         Text            =   "da sistemare ricevi comandi chat"
         Top             =   960
         Width           =   2535
      End
   End
   Begin VB.PictureBox Picture17 
      BackColor       =   &H80000013&
      Height          =   6255
      Left            =   10080
      ScaleHeight     =   6195
      ScaleWidth      =   4875
      TabIndex        =   81
      Top             =   10800
      Width           =   4935
   End
   Begin VB.PictureBox Picture16 
      BackColor       =   &H80000013&
      Height          =   3855
      Left            =   8640
      ScaleHeight     =   3795
      ScaleWidth      =   4635
      TabIndex        =   80
      Top             =   10800
      Width           =   4695
   End
   Begin VB.PictureBox Picture15 
      BackColor       =   &H80000013&
      Height          =   6615
      Left            =   1680
      ScaleHeight     =   6555
      ScaleWidth      =   6075
      TabIndex        =   79
      Top             =   10800
      Width           =   6135
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H80000013&
      Height          =   8775
      Left            =   2760
      ScaleHeight     =   8715
      ScaleWidth      =   6075
      TabIndex        =   78
      Top             =   10800
      Width           =   6135
   End
   Begin VB.PictureBox Picture13 
      BackColor       =   &H80000013&
      Height          =   9735
      Left            =   1320
      ScaleHeight     =   9675
      ScaleWidth      =   6795
      TabIndex        =   77
      Top             =   10800
      Width           =   6855
   End
   Begin VB.PictureBox Picture12 
      BackColor       =   &H80000013&
      Height          =   6135
      Left            =   4560
      ScaleHeight     =   6075
      ScaleWidth      =   3555
      TabIndex        =   76
      Top             =   10800
      Width           =   3615
   End
   Begin VB.PictureBox Picture11 
      BackColor       =   &H80000013&
      Height          =   4215
      Left            =   7920
      ScaleHeight     =   4155
      ScaleWidth      =   6435
      TabIndex        =   75
      Top             =   10800
      Width           =   6495
   End
   Begin VB.PictureBox Picture10 
      BackColor       =   &H80000013&
      Height          =   2775
      Left            =   7440
      ScaleHeight     =   2715
      ScaleWidth      =   4395
      TabIndex        =   74
      Top             =   10800
      Width           =   4455
   End
   Begin VB.PictureBox Picture9 
      BackColor       =   &H8000000E&
      Height          =   1935
      Left            =   6960
      ScaleHeight     =   1875
      ScaleWidth      =   4755
      TabIndex        =   73
      Top             =   10800
      Width           =   4815
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H80000013&
      Height          =   2295
      Left            =   4080
      ScaleHeight     =   2235
      ScaleWidth      =   3915
      TabIndex        =   72
      Top             =   10800
      Width           =   3975
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H80000013&
      Height          =   6255
      Left            =   3600
      ScaleHeight     =   6195
      ScaleWidth      =   6555
      TabIndex        =   71
      Top             =   10800
      Width           =   6615
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H80000013&
      Height          =   5415
      Left            =   1560
      ScaleHeight     =   5355
      ScaleWidth      =   8715
      TabIndex        =   70
      Top             =   10800
      Width           =   8775
   End
   Begin VB.PictureBox Picture5 
      Height          =   4335
      Left            =   7800
      ScaleHeight     =   4275
      ScaleWidth      =   4515
      TabIndex        =   69
      Top             =   11160
      Width           =   4575
   End
   Begin VB.PictureBox Picture4 
      Height          =   6255
      Left            =   1560
      ScaleHeight     =   6195
      ScaleWidth      =   9075
      TabIndex        =   68
      Top             =   11400
      Width           =   9135
   End
   Begin VB.PictureBox Picture3 
      Height          =   2295
      Left            =   480
      ScaleHeight     =   2235
      ScaleWidth      =   5955
      TabIndex        =   67
      Top             =   10560
      Width           =   6015
   End
   Begin VB.PictureBox Picture2 
      Height          =   2895
      Left            =   1440
      ScaleHeight     =   2835
      ScaleWidth      =   4155
      TabIndex        =   66
      Top             =   10680
      Width           =   4215
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   9840
      TabIndex        =   59
      Top             =   10680
      Width           =   2295
      Begin client.CandyButton CandyButton2 
         Height          =   255
         Left            =   1920
         TabIndex        =   60
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "x"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Style           =   6
         Checked         =   0   'False
         ColorButtonHover=   16760976
         ColorButtonUp   =   15309136
         ColorButtonDown =   15309136
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   0
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "MUUUUhHAaHAHAHA"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   65
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "(h5)(h5)(h5)(h5)(h5)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   64
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "XxxxXxxxXxxx "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Left            =   240
         TabIndex        =   63
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "[ :[ : [: [ : [ : [ :["
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   240
         TabIndex        =   62
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Shape Shape4 
         Height          =   2175
         Left            =   0
         Top             =   0
         Width           =   2295
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "LaLalaLAlaLLLAAALLaaa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   61
         Top             =   1800
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   2400
      TabIndex        =   56
      Top             =   10680
      Width           =   4455
      Begin VB.Timer Timer_frame3 
         Enabled         =   0   'False
         Interval        =   3000
         Left            =   3840
         Top             =   120
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   58
         Text            =   "criptaggio con chiave AES 128"
         Top             =   360
         Width           =   2775
      End
      Begin client.chameleonButton chameleonButton2 
         Height          =   1215
         Left            =   0
         TabIndex        =   57
         Top             =   0
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   2143
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "chat.frx":2840C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.CheckBox CheckAES 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   51
      Top             =   10080
      Width           =   255
   End
   Begin VB.CheckBox Checkrisp_veloci 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2520
      TabIndex        =   50
      Top             =   10080
      Width           =   255
   End
   Begin VB.Timer Timer_riabilita_bigsmile 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1320
      Top             =   7920
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000A&
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   49
      Top             =   8040
      Visible         =   0   'False
      Width           =   495
   End
   Begin client.CandyButton Cmdcrediti 
      Height          =   495
      Left            =   7440
      TabIndex        =   48
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "mostra crediti"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   2
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   8454016
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin client.CandyButton Cmdsend 
      Height          =   615
      Left            =   8880
      TabIndex        =   45
      Top             =   9360
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "invia"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   3
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Timer Timer_prepara_uscita_chat 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6960
      Top             =   360
   End
   Begin VB.Timer Timeresci_chat 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   6600
      Top             =   360
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000013&
      Height          =   1335
      Left            =   10680
      TabIndex        =   43
      Top             =   9000
      Visible         =   0   'False
      Width           =   2055
      Begin client.CandyButton Cmdprivileg 
         Height          =   855
         Left            =   360
         TabIndex        =   44
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1508
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "MODERAZIONE"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Style           =   2
         Checked         =   0   'False
         ColorButtonHover=   16760976
         ColorButtonUp   =   15309136
         ColorButtonDown =   15309136
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   0
      End
      Begin VB.Shape Shape3 
         Height          =   1095
         Left            =   240
         Top             =   120
         Width           =   1575
      End
   End
   Begin client.CandyButton Cmdcomandi 
      Height          =   375
      Left            =   11400
      TabIndex        =   42
      Top             =   8400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "invia comandi"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin client.CandyButton Cmdwebcam 
      Height          =   375
      Left            =   11400
      TabIndex        =   41
      Top             =   8040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "webcam chat"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin client.CandyButton Cmdinviafile 
      Height          =   375
      Left            =   10680
      TabIndex        =   40
      Top             =   8880
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "invia file"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin client.CandyButton CmdMS 
      Height          =   375
      Left            =   10680
      TabIndex        =   39
      Top             =   8400
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "MS"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin client.CandyButton CmdPM 
      Height          =   375
      Left            =   10680
      TabIndex        =   38
      Top             =   8040
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "PM"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Timer Timer_caricamento_T9 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   7320
      Top             =   360
   End
   Begin client.CandyButton Cmdmessaggiomassa 
      Height          =   495
      Left            =   9000
      TabIndex        =   37
      Top             =   1080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "messaggio di massa"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   3
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin client.CandyButton CmdPMserverinvio 
      Height          =   495
      Left            =   4800
      TabIndex        =   36
      Top             =   1080
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "invia messaggio al server"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   3
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin client.CandyButton Cmdsmyle 
      Height          =   375
      Left            =   480
      TabIndex        =   35
      ToolTipText     =   "smile semplici"
      Top             =   8400
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "chat.frx":28428
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin client.CandyButton Cmdopaco 
      Height          =   375
      Left            =   3960
      TabIndex        =   34
      Top             =   1800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "opaco"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin client.CandyButton Cmdtrasparenza 
      Height          =   375
      Left            =   2640
      TabIndex        =   33
      Top             =   1800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "trasparenza"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin client.CandyButton Cmdopzioni 
      Height          =   375
      Left            =   1560
      TabIndex        =   32
      Top             =   1800
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "opzioni"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin client.CandyButton Cmdescichat 
      Height          =   375
      Left            =   240
      TabIndex        =   31
      Top             =   1800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "esci chat"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Timer Timer_unload 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   7680
      Top             =   360
   End
   Begin VB.Timer Timer_winsock_close 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   8040
      Top             =   360
   End
   Begin client.CandyButton Cmdbigsmile 
      Height          =   375
      Left            =   1200
      TabIndex        =   30
      ToolTipText     =   "smile grandi"
      Top             =   8400
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "chat.frx":2885A
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.CommandButton Cmdcripta_AES 
      Caption         =   "criptaggio AES"
      Height          =   375
      Left            =   14040
      TabIndex        =   29
      Top             =   9840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Txtpassword 
      Height          =   285
      Left            =   840
      TabIndex        =   28
      Text            =   "superspeed"
      Top             =   10320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox Checkintellisense 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5760
      TabIndex        =   26
      Top             =   9000
      Width           =   255
   End
   Begin VB.CheckBox CheckT9 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7920
      TabIndex        =   25
      Top             =   9000
      Width           =   255
   End
   Begin VB.CommandButton CmdT9 
      Caption         =   "lista T9"
      Height          =   375
      Left            =   14040
      TabIndex        =   24
      Top             =   9480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox picBuffer 
      Height          =   615
      Left            =   13200
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   23
      Top             =   9600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   360
      ScaleHeight     =   675
      ScaleWidth      =   8355
      TabIndex        =   22
      Top             =   9240
      Width           =   8415
      Begin VB.TextBox txtsend 
         Height          =   735
         IMEMode         =   3  'DISABLE
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   0
         Width           =   8295
      End
   End
   Begin client.chameleonButton cmdIM 
      Height          =   975
      Left            =   11160
      TabIndex        =   21
      Top             =   840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "mesaggistica istantanea"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "chat.frx":28D02
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin client.CandyButton Cmdsfondi 
      Height          =   375
      Left            =   8280
      TabIndex        =   20
      ToolTipText     =   "sfondi"
      Top             =   8400
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "chat.frx":28D1E
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin RichTextLib.RichTextBox txtchat 
      Height          =   5655
      Left            =   360
      TabIndex        =   19
      Top             =   2520
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   9975
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"chat.frx":290B8
   End
   Begin client.Anim Anim1 
      Height          =   5655
      Left            =   360
      TabIndex        =   18
      Top             =   2520
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   9975
   End
   Begin client.chameleonButton chameleonButton1 
      Height          =   5895
      Left            =   240
      TabIndex        =   17
      Top             =   2400
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   10398
      BTYPE           =   3
      TX              =   "chameleonButton1"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "chat.frx":2913A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin client.CandyButton Cmdsfonditxtchat 
      Height          =   375
      Left            =   9000
      TabIndex        =   16
      ToolTipText     =   "sfondi testo"
      Top             =   8400
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "chat.frx":29156
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000013&
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   4575
      Begin VB.TextBox Txtmionick 
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Labelmionick 
         BackStyle       =   0  'Transparent
         Caption         =   "il tuo nick in chat e' :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Height          =   855
      Left            =   12840
      TabIndex        =   8
      Top             =   1080
      Width           =   2175
      Begin VB.Image Image1 
         Height          =   240
         Left            =   120
         Picture         =   "chat.frx":294F0
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Labelora 
         BackStyle       =   0  'Transparent
         Caption         =   "ora :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   960
         TabIndex        =   11
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Labeldata 
         BackStyle       =   0  'Transparent
         Caption         =   "data :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Timer Timerora 
      Interval        =   1
      Left            =   8760
      Top             =   360
   End
   Begin VB.Timer Timeruseronline 
      Interval        =   100
      Left            =   8400
      Top             =   360
   End
   Begin MSComDlg.CommonDialog Cmdl 
      Left            =   12000
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frameinformazioni 
      BackColor       =   &H80000013&
      Height          =   3135
      Left            =   12840
      TabIndex        =   2
      Top             =   2040
      Width           =   2175
      Begin VB.TextBox Text2 
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtIpUtente 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.PictureBox Picavatar 
         BackColor       =   &H80000013&
         Height          =   1575
         Left            =   240
         ScaleHeight     =   1515
         ScaleWidth      =   1515
         TabIndex        =   3
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Labelnick 
         BackStyle       =   0  'Transparent
         Caption         =   "nick"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Labeliputente 
         BackStyle       =   0  'Transparent
         Caption         =   "txtiputente"
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000013&
         BackStyle       =   1  'Opaque
         Height          =   1815
         Left            =   120
         Shape           =   5  'Rounded Square
         Top             =   1200
         Width           =   1815
      End
   End
   Begin VB.Frame Frameusers 
      BackColor       =   &H80000013&
      Caption         =   "list users"
      Height          =   5775
      Left            =   10680
      TabIndex        =   0
      Top             =   2040
      Width           =   2055
      Begin VB.ListBox listusers 
         Height          =   4785
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   720
         Width           =   1815
      End
      Begin VB.Image Imageonline 
         Height          =   360
         Left            =   120
         Picture         =   "chat.frx":2A762
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Labelonline 
         BackStyle       =   0  'Transparent
         Caption         =   "Online :"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   600
         TabIndex        =   47
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblclientonline 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001784D5&
         Height          =   255
         Left            =   1320
         TabIndex        =   46
         Top             =   240
         Width           =   375
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4320
      Top             =   8280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "muovi componenti"
      Height          =   255
      Left            =   4440
      TabIndex        =   100
      Top             =   10080
      Width           =   1455
   End
   Begin VB.Image image23 
      Height          =   270
      Index           =   0
      Left            =   8520
      Picture         =   "chat.frx":2AA54
      ToolTipText     =   "memorizza link"
      Top             =   10080
      Width           =   645
   End
   Begin VB.Image image22 
      Height          =   270
      Index           =   1
      Left            =   7680
      Picture         =   "chat.frx":2B476
      ToolTipText     =   "calcolatrice"
      Top             =   10080
      Width           =   645
   End
   Begin VB.Image img_menu 
      Height          =   255
      Index           =   10
      Left            =   3480
      Picture         =   "chat.frx":2BE98
      Top             =   8955
      Width           =   240
   End
   Begin VB.Image image21 
      Height          =   270
      Left            =   1440
      Picture         =   "chat.frx":2C20A
      ToolTipText     =   "profilo utente"
      Top             =   8940
      Width           =   645
   End
   Begin VB.Image image20 
      Height          =   270
      Left            =   600
      Picture         =   "chat.frx":2CC2C
      ToolTipText     =   "risposte fatte"
      Top             =   8940
      Width           =   645
   End
   Begin VB.Image Image11 
      Height          =   270
      Left            =   9360
      Picture         =   "chat.frx":2D64E
      ToolTipText     =   "parole strane"
      Top             =   8955
      Width           =   285
   End
   Begin VB.Image Image12 
      Height          =   285
      Left            =   3840
      Picture         =   "chat.frx":2DAC8
      ToolTipText     =   "mostra eventi"
      Top             =   8925
      Width           =   285
   End
   Begin VB.Image Image13 
      Height          =   285
      Left            =   4320
      Picture         =   "chat.frx":2DF7E
      ToolTipText     =   "cerca utente"
      Top             =   8925
      Width           =   285
   End
   Begin VB.Image Image14 
      Height          =   285
      Left            =   4800
      Picture         =   "chat.frx":2E434
      ToolTipText     =   "lista utenti da bannare"
      Top             =   8925
      Width           =   285
   End
   Begin VB.Image Image15 
      Height          =   285
      Left            =   5280
      Picture         =   "chat.frx":2E8EA
      ToolTipText     =   "aggiungi lista amici"
      Top             =   8925
      Width           =   285
   End
   Begin VB.Image iimage17 
      Height          =   240
      Left            =   7080
      Picture         =   "chat.frx":2EDA0
      ToolTipText     =   "blocca attivita' extrachat"
      Top             =   10080
      Width           =   240
   End
   Begin VB.Image image18 
      Height          =   240
      Left            =   6600
      Picture         =   "chat.frx":2F12A
      ToolTipText     =   "blocca tastiera"
      Top             =   10080
      Width           =   240
   End
   Begin VB.Image image19 
      Height          =   240
      Left            =   6120
      Picture         =   "chat.frx":2F4B4
      ToolTipText     =   "mostra tempo"
      Top             =   10080
      Width           =   240
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "risposte veloci"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   55
      Top             =   10080
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "cripta messaggio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   720
      TabIndex        =   54
      Top             =   10080
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "abilita T9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   53
      Top             =   9000
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ricorda parola scritte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   52
      Top             =   9000
      Width           =   1815
   End
   Begin VB.Image image30 
      Height          =   1530
      Left            =   360
      Picture         =   "chat.frx":2F83E
      Stretch         =   -1  'True
      Top             =   8880
      Width           =   9360
   End
   Begin VB.Image image28 
      Height          =   1545
      Left            =   240
      Picture         =   "chat.frx":34500
      Top             =   8880
      Width           =   195
   End
   Begin VB.Image Image29 
      Height          =   1545
      Left            =   9720
      Picture         =   "chat.frx":3555A
      Top             =   8880
      Width           =   195
   End
   Begin VB.Image Image5 
      Height          =   270
      Left            =   7320
      Picture         =   "chat.frx":365B4
      Top             =   8940
      Width           =   285
   End
   Begin VB.Image Image7 
      Height          =   315
      Left            =   11640
      Picture         =   "chat.frx":36A2E
      Top             =   0
      Width           =   180
   End
   Begin VB.Image Image4 
      Height          =   315
      Left            =   7080
      Picture         =   "chat.frx":36D64
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4680
   End
   Begin VB.Image Image3 
      Height          =   870
      Left            =   0
      Picture         =   "chat.frx":37196
      Top             =   0
      Width           =   7155
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      Height          =   10455
      Left            =   0
      Top             =   0
      Width           =   15135
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H80000002&
      BorderColor     =   &H00BAB6B3&
      Height          =   975
      Left            =   10680
      Top             =   7920
      Width           =   2055
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00BAB6B3&
      X1              =   7800
      X2              =   7800
      Y1              =   1800
      Y2              =   2280
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00BAB6B3&
      X1              =   0
      X2              =   10320
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00BAB6B3&
      X1              =   0
      X2              =   10320
      Y1              =   2280
      Y2              =   2280
   End
End
Attribute VB_Name = "chat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Dim frmResize As New ControlResizer
Dim Newalert As New cpopup ' dichiariamo la variabile per aprire i popup '
                           ' richiamandoli dal modulo cpopup'
Dim numeromessaggi As Integer
Dim numerobigsmile As Integer ' contiamo quante volte viengono usati gli smile grandi'

Private Sub Check_muovi_componenti_Click()
  If Check_muovi_componenti = 1 Then
     Timer_frame6.Enabled = True
     Frame6.Left = 3480
     Frame6.Top = 7440
  End If
End Sub

Private Sub Cmd_animazioni_flash_Click()
 animazioni_flash_chat.Show
End Sub

Private Sub Cmd_animazioni_flash_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Check_muovi_componenti = 1 Then
  movey Button, chat, Cmd_animazioni_flash, False
  movex Button, chat, Cmd_animazioni_flash
 End If
End Sub

Private Sub Cmd_multicanale_Click()
 
End Sub

Private Sub Cmdbold_Click()
 txtsend.Font.Bold = Not txtsend.Font.Bold
End Sub

Private Sub cmdchiudi_frame6_Click()
 Frame6.Left = 480
 Frame6.Top = 10560
End Sub

Private Sub Cmdcolor_Click()
 CommonDialog1.ShowColor
 txtsend.ForeColor = CommonDialog1.Color
End Sub

Private Sub Cmdcolor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Check_muovi_componenti = 1 Then
  movey Button, chat, Cmdcolor, False
  movex Button, chat, Cmdcolor
 End If
End Sub


Private Sub Cmddatabase_immagini_Click()
 seleziona_database_immagine.Show
End Sub

Private Sub Cmditalic_Click()
 txtsend.Font.Italic = Not txtsend.Font.Italic
End Sub

Private Sub Cmdplayer_Click()
 player_audio.Show
End Sub

Private Sub Cmdplayer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Check_muovi_componenti = 1 Then
  movey Button, chat, Cmdplayer, False
  movex Button, chat, Cmdplayer
 End If
End Sub

Private Sub Cmdtxtchat_color_Click()
 CommonDialog1.ShowColor
 txtchat.BackColor = CommonDialog1.Color
End Sub

Private Sub Cmdtxtchat_color_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Check_muovi_componenti = 1 Then
  movey Button, chat, Cmdtxtchat_color, False
  movex Button, chat, Cmdtxtchat_color
 End If
End Sub

Private Sub Cmdtxtsend_color_Click()
 CommonDialog1.ShowColor
 txtsend.BackColor = CommonDialog1.Color
End Sub

Private Sub Cmdtxtsend_color_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Check_muovi_componenti = 1 Then
  movey Button, chat, Cmdtxtsend_color, False
  movex Button, chat, Cmdtxtsend_color
 End If
End Sub

Private Sub Cmdunderline_Click()
 txtsend.Font.Underline = Not txtsend.Font.Underline
End Sub

Private Sub Command1_Click()
 
End Sub

Private Sub Form_Load()

 lblclientonline.Caption = listusers.ListCount ' ci dice quanti users ci sono online in quel momento'
  
 EnableURLDetect txtchat.hwnd, Me.hwnd ' avviamo la funzione che riesce a identificare un link nel txtchat'
                                       ' questa funzione si trova nel modulo hyperlink...thenks sandro, professional programmer '

 Txtmionick.Text = login.Txtnick.Text
 If login.Checkattivasfondi_txtchat = 1 Then
   Call SetTransparent(txtchat.hwnd)
 End If
   'allavvio carichiamo il setparent per la connessione del multicanale'
   SetParent frmClient.hwnd, Picture22.hwnd
    frmClient.Show
    Picture22.Top = 240
    Picture22.Left = 120
    frmClient.Move 0, 0
   frmResize.KeepRatio = True
   frmResize.FontResize = True
   Call frmResize.InitializeResizer(Me)
 End Sub

Private Sub CandyButton2_Click()
 Frame5.Top = 10680
End Sub

Private Sub CheckAES_Click()
 If CheckAES = 1 Then
  Frame3.Top = 8000
  Timer_frame3.Enabled = True
 End If
End Sub

Private Sub Checkintellisense_Click()
If Checkintellisense = 1 Then
 CheckT9 = 0
End If
End Sub

Private Sub Checkrisp_veloci_Click()
 If Checkrisp_veloci = 1 Then
    SetParent risposte_veloci.hwnd, chat.Picture14.hwnd
    risposte_veloci.Show
    Picture14.Top = 1200
    risposte_veloci.Move 0, 0
 ElseIf Checkrisp_veloci = 0 Then
    Picture14.Top = 10800
 End If
End Sub

Private Sub CheckT9_Click()
If CheckT9 = 1 Then
   avviso.Show
   avviso.Labelmessaggio.Caption = "si  sta' caricando la lista del T9, richiedera' qualche secondo"
   Timer_caricamento_T9.Enabled = True ' richiamiamo il timer per caricare la lista nel modo piu' leggero'
ElseIf CheckT9 = 0 Then
   Unload lista_T9
ElseIf CheckT9 = 1 Then
 Checkintellisense = 0
End If
End Sub

Private Sub Cmdcripta_AES_Click()
 Dim AES As New CRijndael
    txtsend = AES.JustCrypter(txtsend, Txtpassword)
End Sub

Private Sub cmdIM_Click()
avviso.Show
avviso.Labelmessaggio.Caption = " non hai abbastanza privilegi"
End Sub

Private Sub Cmdprivileg_Click()
PRIVILEGI.Show
End Sub

Private Sub cmdSend_Click()
Dim utente As String ' dichiariamo una variabile che identifichi tutto il testo da spedire'
 numeromessaggi = numeromessaggi + 1
 sistema_di_crediti.Text1.Text = numeromessaggi
 
  ' sciviamo il codice che ci permettera' di elaborare il testo proima di spedirlo'
  ' in modo da prevenire parolacce e frasi poco buone, per il decoro della chat'
  '                                                                            '
  '                    ringraziamento a PIERINO89 of p2pforum                  '
  ''
  Dim a() As String
a = Split(txtsend.Text, " ")
For X = LBound(a) To UBound(a)
For t = 0 To lista_bannaggioparole.List_bannaggioparole.ListCount
If a(X) = lista_bannaggioparole.List_bannaggioparole.List(t) Then a(X) = ""
Next
Next
'ricomponi la frase
txtsend.Text = ""
For X = LBound(a) To UBound(a)
txtsend.Text = txtsend.Text + " " + a(X)
Next
txtsend.Text = Mid(txtsend.Text, 2)
'-------------------------------------------------------------------------------'
If CheckAES = 1 Then ' cifriamo con chiave AES 128 '
 Cmdcripta_AES_Click
End If

    ' Richiamo la funzione FindAndReplace
    bigsmile.FindAndReplace
 
If login.Txtnick.Text = Txtmionick.Text Then
   utente = Chr(127) & "nick:" & Txtmionick.Text & Chr(127) & "frase:" & login.Txtfrase.Text & Chr(127) & vbCrLf & txtsend.Text
Else
   utente = Chr(127) & "nick:" & login.Txtnick.Text & "> " & " alias " & " <" & Txtmionick.Text & Chr(127) & "frase:" & login.Txtfrase.Text & Chr(127) & vbCrLf & txtsend.Text
End If
If Not txtsend.Text = "" Then
    'txtchat.Text = txtchat.Text + " <" & login.Txtnick.Text & "> " & txtSend.Text + vbCrLf'
    login.Ws.SendData utente ' spediamo la variabile utente che identifica tutta la stringa'
    'Txtsend.Text = "" ' dopo aver spedito il messaggio il txtsend ritorna vuoto'
End If
End Sub

Private Sub Cmdbigsmile_Click()
 SetParent bigsmile.hwnd, chat.Picture4.hwnd
 bigsmile.Show
 Picture4.Top = 1800
 bigsmile.Move 0, 0
 If numerobigsmile < 4 Then numerobigsmile = numerobigsmile + 1
    Text1.Text = numerobigsmile
    Picture4.Top = 1800
 If Text1.Text = "4" Then
    Text1.Text = ""
    Picture4.Visible = False
    Cmdbigsmile.Enabled = False
    Timer_riabilita_bigsmile.Enabled = True
 End If
End Sub

Private Sub Cmdbigsmile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Check_muovi_componenti = 1 Then
  movey Button, chat, Cmdbigsmile, False
  movex Button, chat, Cmdbigsmile
 End If
End Sub

Private Sub Cmdopaco_Click()
MakeOpaque Me.hwnd
End Sub

Private Sub Cmdopzioni_Click()
 SetParent opzioni.hwnd, chat.Picture6.hwnd
 Picture6.Top = 2520
 opzioni.Show
 opzioni.Move 0, 0
End Sub

Private Sub Cmdsfondi_Click()
 SetParent sfondi.hwnd, chat.Picture5.hwnd
 sfondi.Show
 Picture5.Top = 3840
 sfondi.Move 0, 0
End Sub

Private Sub Cmdsfondi_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Check_muovi_componenti = 1 Then
  movey Button, chat, Cmdsfondi, False
  movex Button, chat, Cmdsfondi
 End If
End Sub

Private Sub Cmdsfonditxtchat_Click()
If login.Checkattivasfondi_txtchat = 0 Then
 avviso.Show
 avviso.Labelmessaggio.Caption = " devi attivare dal login gli sfondi per avere accesso a questa opzion "
Else
 SetParent sfondi_txtchat.hwnd, chat.Picture16.hwnd
 sfondi_txtchat.Show
 Picture16.Top = 4200
 sfondi_txtchat.Move 0, 0
End If
End Sub

Private Sub Cmdsfonditxtchat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Check_muovi_componenti = 1 Then
  movey Button, chat, Cmdsfonditxtchat, False
  movex Button, chat, Cmdsfonditxtchat
 End If
End Sub

Private Sub Cmdtrasparenza_Click()
 MakeTransparent Me.hwnd, 200
End Sub

Private Sub Cmdwebcam_Click()
 frmRyCamV2.Show ' questo form per la webcam non e' mio quindi tutti i riconoscimenti vanno al suo autore'
End Sub

Private Sub Cmdmessaggiomassa_Click()
 messaggiomassa.Show
End Sub

Private Sub CmdPMserverinvio_Click()
PMserverinvio.Show
End Sub

Private Sub CmdT9_Click()
Dim sLista As String
Dim sParole() As String
Dim i As Long

Open App.Path & "\listaparole\T9\Lista T9.dat" For Input As #1
Line Input #1, sLista
sParole = Split(sLista, vbLf)
Close #1

For i = LBound(sParole) To UBound(sParole)
lista_T9.ListaT9.AddItem sParole(i)
Next
lista_T9.ListaT9.ListIndex = -1
lista_T9.Show
End Sub

Private Sub Cmdcomandi_Click()
 SetParent invia_comandi_chat.hwnd, chat.Picture17.hwnd
 invia_comandi_chat.Show
 Picture17.Top = 1800
 invia_comandi_chat.Move 0, 0
End Sub

Private Sub Cmdcrediti_Click()
 sistema_di_crediti.Show
End Sub

Private Sub Form_Resize()

  Call frmResize.FormResized(Me)
    
End Sub

Private Sub Cmdinviafile_Click()
FILEinvia.Show ' richiamiamo il form per l'invio dei file'
End Sub

Private Sub CmdMS_Click()
 MSinvio_style.Show
 listusers.Selected(informazioniutente.Txtutente.Text) = False
End Sub

Private Sub CmdPM_Click()
 PMinvio.Show
 listusers.Selected(informazioniutente.Txtutente.Text) = False
End Sub

Private Sub Cmdsmyle_Click()
 SetParent smyle.hwnd, chat.Picture3.hwnd
 smyle.Show
 Picture3.Top = 5880
 smyle.Move 0, 0  ' ho preferito gestire gli smyle su un form separato, vedremo se e' una buona idea '
End Sub

Private Sub Cmdsmyle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Check_muovi_componenti = 1 Then
  'if you comment out movex or movey it will only move left/right or up/down
  movey Button, chat, Cmdsmyle, False
  movex Button, chat, Cmdsmyle
 End If
End Sub

Private Sub iimage17_Click()
 SetParent block_privat.hwnd, chat.Picture11.hwnd
 block_privat.Show
 Picture11.Top = 5760
 block_privat.Move 0, 0
End Sub

Private Sub Image11_Click()
   Frame5.Top = 8500
End Sub

Private Sub Image12_Click()
 SetParent eventi.hwnd, chat.Picture7.hwnd
 eventi.Show
 Picture7.Top = 2640
 eventi.Move 0, 0
End Sub

Private Sub Image13_Click()
 SetParent cercautente.hwnd, chat.Picture8.hwnd
 cercautente.Show
 Picture8.Top = 6600
 cercautente.Move 0, 0
End Sub

Private Sub Image14_Click()
 SetParent listabannati.hwnd, chat.Picture12.hwnd
 listabannati.Show
 Picture12.Top = 2640
 listabannati.Move 0, 0
End Sub

Private Sub Image15_Click()
 listaamici.Show
End Sub

Private Sub image18_Click()
 SetParent blocca_sblocca.hwnd, chat.Picture10.hwnd
 blocca_sblocca.Show
 Picture10.Top = 7200
 blocca_sblocca.Move 0, 0
End Sub

Private Sub image19_Click()
 SetParent tempo.hwnd, chat.Picture9.hwnd
 tempo.Show
 Picture9.Top = 8040
 tempo.Move 0, 0
End Sub

Private Sub image21_Click()
  SetParent profiloutente.hwnd, chat.Picture15.hwnd
  profiloutente.Show
  Picture15.Top = 2280
  profiloutente.Move 0, 0
End Sub

Private Sub image22_Click(Index As Integer)
 SetParent Calcolatrice.hwnd, chat.Picture20.hwnd
 Calcolatrice.Show
 Picture20.Top = 1920
 Calcolatrice.Move 0, 0
End Sub

Private Sub image23_Click(Index As Integer)
 SetParent link.hwnd, chat.Picture21.hwnd
 link.Show
 Picture21.Top = 1920
 link.Move 0, 0
End Sub

Private Sub img_menu_Click(Index As Integer)
 SetParent cambianick.hwnd, chat.Picture19.hwnd
 cambianick.Show
 Picture19.Top = 6600
 cambianick.Move 0, 0
End Sub

Private Sub Label10_Click()
 txtsend.Text = Label10.Caption
 cmdSend_Click
 CandyButton2_Click
End Sub

Private Sub Label11_Click()
 txtsend.Text = Label11.Caption
 cmdSend_Click
 CandyButton2_Click
End Sub

Private Sub Label12_Click()
  txtsend.Text = Label12.Caption
 cmdSend_Click
 CandyButton2_Click
End Sub

Private Sub Label8_Click()
 txtsend.Text = Label8.Caption
 cmdSend_Click
 CandyButton2_Click
End Sub

Private Sub Label9_Click()
 txtsend.Text = Label9.Caption
 cmdSend_Click
 CandyButton2_Click
End Sub



Private Sub Listusers_Click()
' facciamo in modo che si spuntino solo 1 elemento nella lista'
' quando un nuovo elemento verra' selezionato la spunta sul precedente '
' verra' rimossa.....QUESTO E' STATO DIFFICILE'
'                                              '
' CODED BY MUTEBLASTER OF P2PFORUM             '
Dim X ' dichiariamo la variabile che indica gli elementi della lista'
 Dim bSelected As Integer ' quando seleziono uno automaticamente viene ripremuto il precedente cosi' la spunta sparisce '
 For X = 0 To listusers.ListCount - 1
    bSelected = listusers.ListIndex
        If X = bSelected Then
            listusers.Selected(X) = True
        End If
    Next X
 ' cominciamo la separazione dei vari componenti della lista cioe'
 ' nickname , ip ed indice per l'avatar'
 
 ' QUESTA PARTE DI CODICE E' CODED BY ROBY66 OF P2PFORUM '
 informazioniutente.Txtrecord = listusers.Text
 informazioniutente.Txtutente.Text = listusers.ListIndex
 informazioniutente.txtPosizioneSpazi.Text = InStr(informazioniutente.Txtrecord.Text, informazioniutente.txtInput.Text) 'viene calcolata la posizione della virgola del record
 informazioniutente.Txtip.Text = Mid$(informazioniutente.Txtrecord.Text, informazioniutente.txtPosizioneSpazi + 2, Left)
 informazioniutente.Text2.Text = Left$(informazioniutente.Txtrecord.Text, informazioniutente.txtPosizioneSpazi.Text - 1)
 
   informazioniutente.txtPosizioneSpaziIpUtente.Text = InStr(informazioniutente.Txtip.Text, informazioniutente.txtInput.Text) 'viene calcolata la posizione della virgola del record
 informazioniutente.txtIpUtente.Text = Left(informazioniutente.Txtip.Text, informazioniutente.txtPosizioneSpaziIpUtente.Text) 'AAAAAAAAAAHHHHHHHHHHHH!!!!!!!!!!!! (Urlo da stress! puoi cancellarlo)
 
 informazioniutente.txtPosizioneSpaziAvatar.Text = InStr(informazioniutente.Txtip.Text, informazioniutente.txtInput.Text) 'viene calcolata la posizione della virgola del record
 informazioniutente.Txtavataramico.Text = Mid$(informazioniutente.Txtip.Text, informazioniutente.txtPosizioneSpaziAvatar.Text + 3, Left)  'mi mostra il numero dell'Avatar
  ' GRAZIE ROBY PER IL GRANDE AIUTO '
 Picavatar.Picture = LoadPicture(App.Path & "\avatar" & "\immagine" & informazioniutente.Txtavataramico & ".gif") ' richiamiamo l'immagine in base all'indice'
                                                                                                ' che nel form avatar si e' inserito '
 informazioniutente.Picavatar.Picture = LoadPicture(App.Path & "\avatar" & "\immagine" & informazioniutente.Txtavataramico & ".gif")
 Cmdinviafile.Enabled = True
 CmdPM.Enabled = True
 CmdMS.Enabled = True
 Cmdwebcam.Enabled = True
 Call AddHScroll(listusers) ' solo dopo che si e' selezuinato un utente'
                            ' nella liste richiamiamo dal modulo listbox la funzione per'
                            ' avere la barra orizzontale'' BEL TRUCCHETTO QUESTO'
End Sub

Private Sub image20_Click()
  SetParent risposte.hwnd, chat.Picture13.hwnd
  risposte.Show
  Picture13.Top = 240
  risposte.Move 0, 0
End Sub

Private Sub Timer_caricamento_T9_Timer()
CmdT9_Click
Timer_caricamento_T9.Enabled = False
End Sub

' usiamo un timer per far funzionare il setparent anche da un altro forms....'
Private Sub Timer_chiusura_frmrooms_Timer()
 Picture23.Top = 4800
 Picture23.Left = 120
 Unload frmRooms
 Timer_chiusura_frmrooms.Enabled = False
End Sub

Private Sub Timer_frame3_Timer()
 Frame3.Top = 10680
End Sub

Private Sub Timer_frame6_Timer()
 Frame6.Left = 480
 Frame6.Top = 10560
 Timer_frame6.Enabled = False
End Sub

Private Sub Timer_prepara_uscita_chat_Timer()
 'login.Cmddisconnetti.Enabled = True'
 SetParent avviso_chiusura.hwnd, chat.Picture2.hwnd
 avviso_chiusura.Show
 avviso_chiusura.Move 0, 0
 Picture2.Top = 4000
 Picture2.Left = 2000
 login.cmdCreate.Enabled = False
 Timer_prepara_uscita_chat.Enabled = False
End Sub

Private Sub Timer_riabilita_bigsmile_Timer()
 Cmdbigsmile.Enabled = True
 Picture4.Visible = True
 numerobigsmile = 0
 Timer_riabilita_bigsmile.Enabled = False
End Sub

Private Sub Timer_setparent_frmclient_Timer()
 SetParent frmClient.hwnd, Picture22.hwnd
 frmClient.Show
 Picture22.Top = 240
 Picture22.Left = 120
 frmClient.Move 0, 0
 frmClient.Timer_connessione.Enabled = True
 Timer_setparent_frmclient.Enabled = False
End Sub

Private Sub Timer_setparent_frmrooms_Timer()
   SetParent frmRooms.hwnd, Picture23.hwnd
    frmRooms.Show
    frmRooms.Move 0, 0
    Picture23.Top = 240
    Picture23.Left = 120
    Timer_setparent_frmrooms.Enabled = False
End Sub

Private Sub Timer_unload_Timer()
login.Visible = True
chat.Visible = False
chat_style.Visible = False
Picture2.Left = 1440
Picture2.Top = 10680
Timer_unload.Enabled = False
End Sub

Private Sub Timer_winsock_close_Timer()
' mettiamo un timer che regoli la uscita per dare respiro al sistema'
' infatti l'uscita e' ricca di avvenimenti'
login.WsMSricevi.Close
login.WsPMricevi.Close
login.Wsricevifile.Close
login.WsPMserverricevi.Close
login.Wsricevicomandichat.Close
Timer_winsock_close.Enabled = False
End Sub

Private Sub Timeresci_chat_Timer()
 login.Ws.SendData "@DISCONNECT:" & login.Txtnick.Text & "   " & login.Label1 & "   " & avatar.Txtavatar.Text
 DisableURLDetect
 Timeresci_chat.Enabled = False
End Sub

Private Sub Timerora_Timer()
Label1.Caption = Time
Label2.Caption = Date
End Sub

' dichiariamo un timer che verifichi ogni tanto quanti user sono in linea '
Private Sub Timeruseronline_Timer()
Timeruseronline.interval = 100
DoEvents
lblclientonline = listusers.ListCount
End Sub

Private Sub txtchat_Change()
    ' Richiamo la funzione FindAndReplace per 5 volte (dovrebbero bastare per visualizzare 5 smile)
    For X = 0 To 5
        bigsmile.FindAndReplace
    Next X
End Sub

Private Sub txtSend_Change()
 ' blocchiamo possibili attacchi per salire di grado nel punteggio.....ogni volta che si preme il bottone'
 ' il punteggio aumenta, per prevenire che si schiacci 1000 volte un bottone a vuoto'
 ' se il txtsend e' vuoto il bottone si disabilita'
 If txtsend.Text = "" Then
    Cmdsend.Enabled = False
 Else
    Cmdsend.Enabled = True
 End If
'SostituisciSmyle Txtsend.Text
If Checkintellisense = 1 Then
   iSenseChange txtsend
End If
If CheckT9 = 1 Then ' SE L'UTENTE ABILITA IL T9 SI PROCEDE CON L'EVENTO'
    If KeyPressed = 8 Or KeyPressed = vbKeyDelete Then
        If txtsend.Text = "" Then
            lista_T9.ListaT9.ListIndex = -1
        Else
        End If
    Else
        Dim i As Integer
        Dim strEntry As String
        Dim strStored As String
        Dim placeholder As Integer
                
        strEntry = txtsend.Text
    
        If strEntry <> "" Then
            For i = 0 To lista_T9.ListaT9.ListCount - 1
                
                strStored = lista_T9.ListaT9.List(i)
                If LCase(strEntry) = LCase(Left(strStored, Len(strEntry))) Then
                    txtsend.Text = strStored
                    txtsend.SelStart = Len(strEntry)
                    txtsend.SelLength = Len(strStored) - Len(strEntry)
                    Exit For
                End If
            Next
                placeholder = lista_T9.ListaT9.ListIndex
                
                If i = lista_T9.ListaT9.ListCount Then
                    lista_T9.ListaT9.ListIndex = placeholder
                Else
                    lista_T9.ListaT9.ListIndex = i
                End If
        End If
    End If
  End If
End Sub


Private Sub txtSend_KeyPress(KeyAscii As Integer)
If Checkintellisense = 1 Then
   iSenseKeyPress txtsend, KeyAscii
End If
If KeyAscii = 13 Then ' se si preme enter'
    KeyAscii = 0
    cmdSend_Click
End If
End Sub

Private Sub txtSend_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyPressed = KeyCode
End Sub

 ' aggiungiamo un bottone per la disconnessione del programma, che permettera' una corrtetta uscita'
Private Sub Cmdescichat_Click()
 Timer_prepara_uscita_chat.Enabled = True
 Timeresci_chat.Enabled = True
 Timer_unload.Enabled = True
End Sub
