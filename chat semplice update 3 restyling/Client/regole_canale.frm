VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form regole_canale 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "regole canale di chat e messenger"
   ClientHeight    =   8325
   ClientLeft      =   1095
   ClientTop       =   1290
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   6405
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000013&
      Height          =   8295
      Left            =   0
      ScaleHeight     =   8235
      ScaleWidth      =   6315
      TabIndex        =   48
      Top             =   8640
      Width           =   6375
   End
   Begin client.CandyButton Cannulla 
      Height          =   375
      Left            =   4440
      TabIndex        =   47
      Top             =   7800
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "annulla registrazione"
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
   Begin client.CandyButton Ccmdchiudi 
      Height          =   495
      Left            =   5760
      TabIndex        =   46
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "X"
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
   Begin VB.Frame Frame7 
      BorderStyle     =   0  'None
      Height          =   8295
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   6375
      Begin client.CandyButton CandyButton7 
         Height          =   375
         Left            =   600
         TabIndex        =   45
         Top             =   6720
         Width           =   1695
         _ExtentX        =   2990
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
         Caption         =   "<<<<< indietro"
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
      Begin client.CandyButton CandyButton6 
         Height          =   375
         Left            =   3600
         TabIndex        =   44
         Top             =   6720
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
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
         Caption         =   "avantii  >>>>>"
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
      Begin VB.CheckBox Check5 
         Caption         =   "dichiaro di non avere ancora raggiunto la maggior eta' "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1200
         TabIndex        =   42
         Top             =   4320
         Width           =   2775
      End
      Begin VB.CheckBox Check4 
         Caption         =   "dichiaro di aver raggiunto la maggior eta'"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1200
         TabIndex        =   41
         Top             =   2640
         Width           =   3135
      End
      Begin VB.Label Label5 
         Caption         =   "per completare la registrazione , e' opportuno di chiarare se si e' maggiorenni oppure no"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   120
         TabIndex        =   43
         Top             =   480
         Width           =   5055
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   8295
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   6375
      Begin client.Anim Anim1 
         Height          =   1575
         Left            =   4080
         TabIndex        =   25
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   2778
      End
      Begin client.chameleonButton chameleonButton4 
         Height          =   1815
         Left            =   3960
         TabIndex        =   24
         Top             =   120
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   3201
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
         BCOL            =   65280
         BCOLO           =   16711935
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "regole_canale.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4545
         ItemData        =   "regole_canale.frx":001C
         Left            =   480
         List            =   "regole_canale.frx":00C8
         TabIndex        =   19
         Top             =   2160
         Width           =   5415
      End
      Begin client.chameleonButton chameleonButton1 
         Height          =   5055
         Left            =   120
         TabIndex        =   18
         Top             =   1920
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   8916
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
         MICON           =   "regole_canale.frx":0E8E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.CheckBox Check_preso_coscenza 
         Caption         =   "si ho preso coscenza dei consigli della polizia postale"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   10
         Top             =   7080
         Width           =   2175
      End
      Begin VB.CheckBox Checknon_preso_coscenza 
         Caption         =   "non prendo coscenza di questi consigli"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         TabIndex        =   9
         Top             =   7200
         Width           =   2295
      End
      Begin client.CandyButton Critornaindietro1 
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   7800
         Width           =   1575
         _ExtentX        =   2778
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
         Caption         =   "<<<<<indietro"
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
      Begin client.CandyButton Cmdregistrati 
         Height          =   375
         Left            =   3600
         TabIndex        =   8
         Top             =   7800
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
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
         Caption         =   "prosegui>>>>>"
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "LEGGERE ATTENTAMENTE"
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
         Left            =   360
         TabIndex        =   12
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "riporto un articolo di avvertenze che la polizia postele ha redatto apropposito delle chat "
         Height          =   615
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   8295
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   6375
      Begin VB.Frame Frame6 
         Height          =   2535
         Left            =   1320
         TabIndex        =   35
         Top             =   1920
         Visible         =   0   'False
         Width           =   3855
         Begin client.CandyButton CandyButton5 
            Height          =   375
            Left            =   1920
            TabIndex        =   39
            Top             =   1440
            Width           =   1575
            _ExtentX        =   2778
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
            Caption         =   "rifiuto"
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
         Begin client.CandyButton CandyButton4 
            Height          =   375
            Left            =   120
            TabIndex        =   38
            Top             =   1440
            Width           =   1455
            _ExtentX        =   2566
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
            Caption         =   "accetto "
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
         Begin RichTextLib.RichTextBox RichTextBox2 
            Height          =   855
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   1508
            _Version        =   393217
            BackColor       =   -2147483645
            ReadOnly        =   -1  'True
            TextRTF         =   $"regole_canale.frx":0EAA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin client.chameleonButton chameleonButton6 
            Height          =   2535
            Left            =   0
            TabIndex        =   36
            Top             =   0
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   4471
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
            MICON           =   "regole_canale.frx":0F79
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
      Begin VB.Frame Frame5 
         Height          =   2535
         Left            =   1320
         TabIndex        =   30
         Top             =   1920
         Visible         =   0   'False
         Width           =   3855
         Begin client.CandyButton CandyButton2 
            Height          =   375
            Left            =   2160
            TabIndex        =   34
            Top             =   1560
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "rifiuto"
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
         Begin client.CandyButton CandyButton1 
            Height          =   375
            Left            =   360
            TabIndex        =   33
            Top             =   1560
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "accetto"
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
         Begin RichTextLib.RichTextBox RichTextBox1 
            Height          =   615
            Left            =   240
            TabIndex        =   32
            Top             =   240
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   1085
            _Version        =   393217
            TextRTF         =   $"regole_canale.frx":0F95
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin client.CandyButton CandyButton3 
            Height          =   2535
            Left            =   0
            TabIndex        =   31
            Top             =   0
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   4471
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
            Checked         =   0   'False
            ColorButtonHover=   16760976
            ColorButtonUp   =   15309136
            ColorButtonDown =   15309136
            BorderBrightness=   0
            ColorBright     =   16772528
            DisplayHand     =   0   'False
            ColorScheme     =   0
         End
      End
      Begin VB.Frame Frame4 
         Height          =   855
         Left            =   3360
         TabIndex        =   28
         Top             =   7200
         Width           =   2775
         Begin VB.CheckBox Check3 
            Caption         =   "esamina licenza in italiano"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   2535
         End
      End
      Begin client.chameleonButton chameleonButton5 
         Height          =   1095
         Left            =   3240
         TabIndex        =   27
         Top             =   7080
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   1931
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
         MICON           =   "regole_canale.frx":104B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ListBox List3 
         Height          =   5520
         ItemData        =   "regole_canale.frx":1067
         Left            =   600
         List            =   "regole_canale.frx":141C
         TabIndex        =   26
         Top             =   1080
         Visible         =   0   'False
         Width           =   5295
      End
      Begin VB.ListBox ListGNU 
         BackColor       =   &H8000000A&
         Height          =   5520
         ItemData        =   "regole_canale.frx":5BF0
         Left            =   600
         List            =   "regole_canale.frx":5F3F
         TabIndex        =   21
         Top             =   1080
         Width           =   5295
      End
      Begin client.chameleonButton chameleonButton2 
         Height          =   6135
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   10821
         BTYPE           =   3
         TX              =   "chameleonButton2"
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
         MICON           =   "regole_canale.frx":9B7F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin client.CandyButton Cmdindietro 
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   7680
         Width           =   1335
         _ExtentX        =   2355
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
         Caption         =   "<<  indietro"
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
      Begin client.CandyButton Cmdsucessivo 
         Height          =   375
         Left            =   1680
         TabIndex        =   16
         Top             =   7680
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
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
         Caption         =   "avanti >>"
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
      Begin VB.CheckBox Check2 
         Caption         =   "rifiuto"
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
         Left            =   2040
         TabIndex        =   15
         Top             =   7200
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "accetto"
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
         TabIndex        =   14
         Top             =   7200
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "l'uso di questo programma e' fatto nel rispetto della licenza GNU"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   480
         TabIndex        =   13
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.ListBox List1 
         Height          =   5715
         ItemData        =   "regole_canale.frx":9B9B
         Left            =   480
         List            =   "regole_canale.frx":9C2F
         TabIndex        =   23
         Top             =   960
         Width           =   5535
      End
      Begin client.chameleonButton chameleonButton3 
         Height          =   6255
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   11033
         BTYPE           =   3
         TX              =   "chameleonButton3"
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
         MICON           =   "regole_canale.frx":A85B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.CheckBox Checkaccetto 
         Caption         =   "accetto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   7080
         Width           =   975
      End
      Begin client.CandyButton Cmdprosegui 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   7680
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "prosegui >>>>>"
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
      Begin VB.CheckBox Checkrifiuto 
         Caption         =   "non accetto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   7080
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "PER ISCRIVERSI E' NECESSARIO LEGGERE LE REGOLE DEL CANALE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   720
         TabIndex        =   4
         Top             =   120
         Width           =   4695
      End
   End
End
Attribute VB_Name = "regole_canale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private Sub CandyButton1_Click()
List3.Visible = True
Frame5.Visible = False
Check1.Enabled = True
End Sub

Private Sub CandyButton2_Click()
Check3 = 0
Frame5.Visible = False
Check1.Enabled = True
End Sub

Private Sub CandyButton4_Click()
 Cmdsucessivo.Visible = True
 Frame6.Visible = False
 Check2.Enabled = True
 Check3 = 0
End Sub

Private Sub CandyButton5_Click()
Check1 = 0
Frame6.Visible = False
Check2.Enabled = True
End Sub

Private Sub CandyButton6_Click()
 SetParent frmCreate.hwnd, Picture1.hwnd
 frmCreate.Show
 Picture1.Top = 0
 frmCreate.Move 0, 0
End Sub

Private Sub CandyButton7_Click()
Frame7.Visible = False
End Sub

Private Sub Cannulla_Click()
 Unload regole_canale
End Sub

Private Sub Ccmdchiudi_Click()
 Cannulla_Click
End Sub

Private Sub Check_preso_coscenza_Click()
If Check_preso_coscenza = 1 Then
   Cmdregistrati.Visible = True
   Checknon_preso_coscenza = 0
ElseIf Check_preso_coscenza = 0 Then
   Cmdregistrati.Visible = False
End If
End Sub

Private Sub Check1_Click()
If Check1 = 1 Then
   Check2 = 0
   Cmdsucessivo.Visible = True
   If List3.Visible = True Then
      Frame6.Visible = True
      Cmdsucessivo.Visible = False
      Check2.Enabled = False
   End If
ElseIf Check1 = 0 Then
   Cmdsucessivo.Visible = False
End If
End Sub

Private Sub Check2_Click()
If Check2 = 1 Then
   Check1 = 0
   Cmdsucessivo.Visible = False
End If
End Sub

Private Sub Check3_Click()
If Check3 = 1 Then
   Frame5.Visible = True
   Check1.Enabled = False
   ElseIf Check3 = 0 Then
   List3.Visible = False
   
End If
End Sub

Private Sub Check4_Click()
If Check4 = 1 Then
 Check5 = 0
 CandyButton6.Visible = True
End If
End Sub

Private Sub Check5_Click()
If Check5 = 1 Then
   Check4 = 0
   CandyButton6.Visible = True
End If
End Sub

Private Sub Checkaccetto_Click()
If Checkaccetto = 1 Then
   Checkrifiuto = 0
   Cmdprosegui.Visible = True
ElseIf Checkaccetto = 0 Then
   Cmdprosegui.Visible = False
End If

End Sub

Private Sub Checknon_preso_coscenza_Click()
If Checknon_preso_coscenza = 1 Then
   Check_preso_coscenza = 0
   Cmdprosegui.Visible = False
End If
End Sub

Private Sub Checkrifiuto_Click()
 If Checkrifiuto = 1 Then
    Checkaccetto = 0
    Cmdprosegui.Visible = False
 End If
End Sub

Private Sub Cmdindietro_Click()
Frame2.Visible = False
End Sub

Private Sub Cmdprosegui_Click()
 Frame2.Visible = True
End Sub

Private Sub Cmdregistrati_Click()
Frame7.Visible = True
End Sub

Private Sub Cmdsucessivo_Click()
Frame3.Visible = 3
End Sub

Private Sub Critornaindietro1_Click()
 Frame3.Visible = False
End Sub

Private Sub Form_Load()
 Call AddHScroll(List1)
 Call AddHScroll(List2)
 Call AddHScroll(ListGNU)
 Call AddHScroll(List3)
 Anim1.AnimatedGifPath = App.Path & "\immagini varie" & "\immagine1.jpg"
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Form_Unload(cancel As Integer)
 Cannulla_Click
End Sub
