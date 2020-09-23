VERSION 5.00
Begin VB.Form block_privat 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "blocco attivita' extra chat"
   ClientHeight    =   4125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin client.CandyButton Cmdexit 
      Height          =   255
      Left            =   5640
      TabIndex        =   9
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
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin client.CandyButton Cmdok 
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      Top             =   3600
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
      Caption         =   "ok"
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
   Begin client.CandyButton Cmdriabilita 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3480
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
      Caption         =   "riabilita"
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
   Begin client.CandyButton CmsRICEVIFILE 
      Height          =   495
      Left            =   3960
      TabIndex        =   6
      Top             =   1920
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "blocca ricezione file"
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
   Begin client.CandyButton CmdblockPM 
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   1920
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "blocca PM"
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
   Begin client.CandyButton CmdbloccoMS 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "blocca MS"
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
   Begin VB.Timer Timerunload 
      Enabled         =   0   'False
      Interval        =   50000
      Left            =   1200
      Top             =   3480
   End
   Begin VB.Shape Shape1 
      Height          =   4095
      Left            =   0
      Top             =   0
      Width           =   6375
   End
   Begin VB.Image Image9 
      Height          =   450
      Left            =   0
      Picture         =   "block_privat.frx":0000
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image11 
      Height          =   435
      Left            =   6000
      Picture         =   "block_privat.frx":08B2
      Top             =   15
      Width           =   300
   End
   Begin VB.Image Image10 
      Height          =   450
      Left            =   240
      Picture         =   "block_privat.frx":0FC0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5805
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "non  riceveri piu' file dagli altri utenti"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "non riceverai piu'  singoli messaggi privati"
      Height          =   615
      Left            =   2160
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "non riceverai piu' chat private"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Labelblocco 
      BackStyle       =   0  'Transparent
      Caption         =   "questi comandi consentono di bloccare tutte le attivita' extrachat .....i messaggi privati, i singoli messaggi ecc......"
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
      TabIndex        =   0
      Top             =   960
      Width           =   5775
   End
End
Attribute VB_Name = "block_privat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' in questo form vengono bloccate tutte quelle operazioni extra chat '
' ho inserito questo form perche' non tutti potrebbero gradire queste opzioni'
' infatti non tutti potrebbero gradire i messaggi provati o le chat private'

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Private OldX As Integer
Private OldY As Integer

Private Sub Cmdexit_Click()
 Cmdok_Click
End Sub

Private Sub Cmdriabilita_Click()
riabilita_privat.Show
Cmdok_Click
End Sub

'all'avvio del programma viene abilitato il timer per far sparire il form'
' e gli viene dato un intervallo di 50 secondi ( 50000 millisecondi)'
Private Sub form_load()
 Timerunload = True
End Sub

Private Sub CmdblockPM_Click()
On Error Resume Next
login.WsPMricevi.Close
End Sub

Private Sub CmsRICEVIFILE_Click()
On Error Resume Next
login.Wsricevifile.Close
End Sub

Private Sub CmdbloccoMS_Click()
On Error Resume Next
login.WsMSricevi.Close
End Sub

Private Sub Cmdok_Click()
 chat.Picture11.Top = 10800
 Unload block_privat
End Sub

Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ReleaseCapture
SendMessage block_privat.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub

' ho inserito un timer in modo che se il form viene dimenticato aperto'
' venga chiuso automaticamente'
Private Sub Timerunload_Timer()
 If block_privat.Visible = True Then
  Cmdok_Click
 End If
End Sub
