VERSION 5.00
Begin VB.Form informazioni_porte 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6390
   ClientLeft      =   4125
   ClientTop       =   1005
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   Begin client.CandyButton Cmdchiudi 
      Height          =   495
      Left            =   2880
      TabIndex        =   4
      Top             =   5400
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "chiudi"
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
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   840
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "informazioni_porte.frx":0000
      Top             =   2160
      Width           =   3015
   End
   Begin client.chameleonButton chameleonButton1 
      Height          =   4335
      Left            =   480
      TabIndex        =   2
      Top             =   1800
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   6588
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
      MICON           =   "informazioni_porte.frx":005C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin client.CandyButton Cmdexit 
      Height          =   255
      Left            =   4200
      TabIndex        =   0
      Top             =   120
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
      Caption         =   "X"
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
   Begin VB.Shape Shape2 
      Height          =   1095
      Left            =   480
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "informazioni sulle porte utilizzate dal programma, per chi ha un firewall molto buono che blocca l'accesso o per chi ha un router"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      Height          =   6375
      Left            =   0
      Top             =   0
      Width           =   4815
   End
   Begin VB.Image Image9 
      Height          =   450
      Left            =   0
      Picture         =   "informazioni_porte.frx":0078
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image11 
      Height          =   435
      Left            =   4440
      Picture         =   "informazioni_porte.frx":092A
      Top             =   15
      Width           =   300
   End
   Begin VB.Image Image10 
      Height          =   450
      Left            =   240
      Picture         =   "informazioni_porte.frx":1038
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4245
   End
End
Attribute VB_Name = "informazioni_porte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Private OldX As Integer
Private OldY As Integer

Private Sub Cmdchiudi_Click()
 Cmdexit_Click
End Sub

Private Sub Cmdexit_Click()
 login.Picture8.Top = 11000
 login.Picture8.Left = 240
End Sub

