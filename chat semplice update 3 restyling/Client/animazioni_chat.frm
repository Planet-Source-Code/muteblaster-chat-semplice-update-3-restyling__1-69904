VERSION 5.00
Begin VB.Form animazioni_chat 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2820
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   2820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin client.Anim Anim1 
      Height          =   2055
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   3625
   End
   Begin client.chameleonButton chameleonButton1 
      Height          =   2535
      Left            =   0
      TabIndex        =   5
      Top             =   480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   4471
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
      MICON           =   "animazioni_chat.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin client.CandyButton Cmdchiudi 
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
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
   Begin VB.TextBox Txtanimazione 
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Top             =   3000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin client.CandyButton Cmdsucessivo 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   3000
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
      Caption         =   ">>>>>"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   6
      Checked         =   0   'False
      ColorButtonHover=   15309136
      ColorButtonUp   =   13657888
      ColorButtonDown =   10512144
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   1
   End
   Begin client.CandyButton Cmdprecedente 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3000
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
      Caption         =   "<<<<<"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   6
      Checked         =   0   'False
      ColorButtonHover=   15309136
      ColorButtonUp   =   13657888
      ColorButtonDown =   10512144
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   1
   End
   Begin client.CandyButton Cmdanim1 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   3000
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
      Caption         =   "ok"
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
   Begin VB.Image Image10 
      Height          =   450
      Left            =   240
      Picture         =   "animazioni_chat.frx":001C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2325
   End
   Begin VB.Image Image11 
      Height          =   435
      Left            =   2520
      Picture         =   "animazioni_chat.frx":0FD6
      Top             =   15
      Width           =   300
   End
   Begin VB.Image Image9 
      Height          =   450
      Left            =   0
      Picture         =   "animazioni_chat.frx":16E4
      Top             =   0
      Width           =   345
   End
End
Attribute VB_Name = "animazioni_chat"
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

Dim immaginenumero As Integer ' dichiariamo la variabile che identifica il numero delle immagini'

Private Sub Cmdanim1_Click()
 esegui_animazioni_chat_inviate.Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\chat" & "\immagine" & immaginenumero & ".gif"
 esegui_animazioni_chat_inviate.Show
 invia_comandi_chat.ws_invia_comandi_chat.SendData Txtanimazione.Text
End Sub

Private Sub Cmdchiudi_Click()
 invia_comandi_chat.Picture1.Top = 5880
End Sub

Private Sub Cmdprecedente_Click()
If immaginenumero > 1 Then immaginenumero = immaginenumero - 1 ' bisogna indicare il numero massimo di gif presenti'
                                                               ' altrimenti continuando a proseguire va' in crash'
Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\chat" & "\immagine" & immaginenumero & ".gif"
Txtanimazione.Text = "animazione" & immaginenumero
End Sub

Private Sub Cmdsucessivo_Click()
If immaginenumero < 20 Then immaginenumero = immaginenumero + 1
  Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\chat" & "\immagine" & immaginenumero & ".gif"
  Txtanimazione.Text = "animazione" & immaginenumero
End Sub

Private Sub form_load()
 Cmdsucessivo_Click
End Sub

Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ReleaseCapture
SendMessage animazioni_chat.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub
