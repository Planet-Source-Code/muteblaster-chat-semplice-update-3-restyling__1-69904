VERSION 5.00
Begin VB.Form PRIVILEGI 
   BackColor       =   &H80000013&
   Caption         =   "privilegi"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   4890
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Height          =   3255
      Left            =   360
      TabIndex        =   6
      Top             =   3240
      Width           =   4095
      Begin VB.Frame Frame3 
         BackColor       =   &H80000013&
         Height          =   1215
         Left            =   120
         TabIndex        =   10
         Top             =   1920
         Width           =   3855
         Begin client.CandyButton CandyButton2 
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   240
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
            Caption         =   "chat"
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
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "chat tra mod"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   720
            Width           =   975
         End
      End
      Begin client.CandyButton CandyButton1 
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   960
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
         Caption         =   "modera"
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
      Begin VB.Shape Shape1 
         Height          =   975
         Left            =   120
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "contatta l'utente da moderare"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "pannello di controllo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   7
         Top             =   120
         Width           =   2535
      End
   End
   Begin client.chameleonButton chameleonButton2 
      Height          =   3855
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   6800
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
      MICON           =   "PRIVILEGI.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.PictureBox Picture3 
         Height          =   375
         Left            =   720
         Picture         =   "PRIVILEGI.frx":001C
         ScaleHeight     =   315
         ScaleWidth      =   1635
         TabIndex        =   4
         Top             =   1920
         Width           =   1695
      End
      Begin VB.PictureBox Picture1 
         Height          =   1695
         Left            =   600
         Picture         =   "PRIVILEGI.frx":06E8
         ScaleHeight     =   1635
         ScaleWidth      =   1995
         TabIndex        =   2
         Top             =   120
         Width           =   2055
         Begin VB.PictureBox Picture2 
            Height          =   15
            Left            =   0
            ScaleHeight     =   15
            ScaleWidth      =   855
            TabIndex        =   3
            Top             =   1680
            Width           =   855
         End
      End
      Begin client.chameleonButton chameleonButton1 
         Height          =   2535
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   3135
         _ExtentX        =   5530
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
         MICON           =   "PRIVILEGI.frx":18AA
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
End
Attribute VB_Name = "PRIVILEGI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CandyButton1_Click()
 MODERAZIONE.Show
End Sub

Private Sub CandyButton2_Click()
 MOD_login.Show
End Sub
