VERSION 5.00
Begin VB.Form psw_sicurezza 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   ClientHeight    =   2670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   ScaleHeight     =   2670
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Height          =   2175
      Left            =   720
      TabIndex        =   4
      Top             =   120
      Width           =   4095
      Begin client.CandyButton Cmdindietro 
         Height          =   495
         Left            =   2760
         TabIndex        =   8
         Top             =   1320
         Width           =   1095
         _ExtentX        =   1931
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
         Caption         =   "rifiuto"
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
      Begin client.CandyButton Cmdprosegui 
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
         _ExtentX        =   2143
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
         Caption         =   "accetto"
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
      Begin VB.TextBox Text1 
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
         Height          =   975
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "password_sicurezza.frx":0000
         Top             =   120
         Width           =   3615
      End
      Begin client.chameleonButton chameleonButton2 
         Height          =   2175
         Left            =   0
         TabIndex        =   5
         Top             =   120
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   3836
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
         MICON           =   "password_sicurezza.frx":006A
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
   Begin client.CandyButton Ccmdsalva 
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   1920
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
      Caption         =   "salva"
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
   Begin VB.TextBox Txtpassword_sicurezza 
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Text            =   "superspeed"
      Top             =   1200
      Width           =   3255
   End
   Begin client.chameleonButton chameleonButton1 
      Height          =   855
      Left            =   720
      TabIndex        =   1
      Top             =   960
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1508
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
      MICON           =   "password_sicurezza.frx":0086
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Shape Shape1 
      Height          =   2655
      Left            =   0
      Top             =   0
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "scrivi la password di sicurezza, con cui vuoi proteggere l'acceso al programma"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "psw_sicurezza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 Txtpassword_sicurezza = RegLoad(Txtpassword_sicurezza)
End Sub
Private Sub Ccmdsalva_Click()
Call SaveControlValues
recupero_password_sicurezza.Show
End Sub

Private Sub Cmdindietro_Click()
Unload psw_sicurezza
End Sub

Private Sub Cmdprosegui_Click()
Frame1.Visible = False
End Sub

Private Sub SaveControlValues()
Call RegSave(Txtpassword_sicurezza, Txtpassword_sicurezza.Text)
End Sub
