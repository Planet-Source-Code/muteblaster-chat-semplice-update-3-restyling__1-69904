VERSION 5.00
Begin VB.Form sfondi 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "sfondi chat"
   ClientHeight    =   4185
   ClientLeft      =   2565
   ClientTop       =   6060
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   Begin client.CandyButton Cmdcondividi_sfondo 
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
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
      Caption         =   "condividi sfondo"
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
   Begin client.CandyButton Cmdanim1 
      Height          =   375
      Left            =   3240
      TabIndex        =   7
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
      Caption         =   "ok"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   16744576
      ColorButtonUp   =   16711680
      ColorButtonDown =   65280
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.TextBox Txtsfondo 
      Height          =   285
      Left            =   3120
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin client.CandyButton Cmdsucessivo 
      Height          =   375
      Left            =   1920
      TabIndex        =   5
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
      Caption         =   ">>>>>"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   4
      Checked         =   0   'False
      ColorButtonHover=   32768
      ColorButtonUp   =   15309136
      ColorButtonDown =   65535
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin client.CandyButton Cmdprecedente 
      Height          =   375
      Left            =   240
      TabIndex        =   4
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
      Caption         =   "<<<<<"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   4
      Checked         =   0   'False
      ColorButtonHover=   16711680
      ColorButtonUp   =   15309136
      ColorButtonDown =   8454143
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin client.Anim Anim1 
      Height          =   2055
      Left            =   480
      TabIndex        =   3
      Top             =   960
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   3625
   End
   Begin client.chameleonButton chameleonButton1 
      Height          =   2535
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   2535
      _ExtentX        =   4471
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
      MICON           =   "sfondi.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CommandButton Cmdchiudi 
      Caption         =   "chiudi"
      Height          =   255
      Left            =   3720
      TabIndex        =   0
      Top             =   3960
      Visible         =   0   'False
      Width           =   615
   End
   Begin client.CandyButton Cmdexit 
      Height          =   255
      Left            =   3720
      TabIndex        =   1
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
      ColorButtonDown =   16711680
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Label Labelcaption_form 
      BackStyle       =   0  'Transparent
      Caption         =   "sfondi per la chat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image Image10 
      Height          =   450
      Left            =   240
      Picture         =   "sfondi.frx":001C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3885
   End
   Begin VB.Image Image11 
      Height          =   435
      Left            =   4080
      Picture         =   "sfondi.frx":0FD6
      Top             =   15
      Width           =   300
   End
   Begin VB.Image Image9 
      Height          =   450
      Left            =   0
      Picture         =   "sfondi.frx":16E4
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image8 
      Height          =   3705
      Left            =   7320
      Picture         =   "sfondi.frx":1F96
      Top             =   600
      Width           =   4440
   End
   Begin VB.Image Image6 
      Height          =   3705
      Left            =   0
      Picture         =   "sfondi.frx":2CE3
      Top             =   480
      Width           =   4440
   End
End
Attribute VB_Name = "sfondi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim immaginenumero As Integer ' dichiariamo la variabile che identifica il numero delle immagini'
Dim frmResize As New ControlResizer

Private Sub Cmdanim1_Click()
 chat_style.Anim1.AnimatedGifPath = App.Path & "\sfondi" & "\chat" & "\immagine" & immaginenumero & ".jpg"
 MakeTransparent chat.hWnd, 200
 Cmdchiudi_Click
End Sub

Private Sub Cmdchiudi_Click()
 sfondi.Visible = False
 chat.Picture5.Top = 11160
End Sub

Private Sub Cmdcondividi_sfondo_Click()
invia_comandi_chat.ws_invia_comandi_chat.SendData Txtsfondo.Text
End Sub

Private Sub Cmdexit_Click()
Cmdchiudi_Click
End Sub

Private Sub Cmdprecedente_Click()
If immaginenumero > 1 Then immaginenumero = immaginenumero - 1 ' bisogna indicare il numero massimo di gif presenti'
                                                               ' altrimenti continuando a proseguire va' in crash'
Anim1.AnimatedGifPath = App.Path & "\sfondi" & "\chat" & "\immagine" & immaginenumero & ".jpg"
Txtsfondo.Text = "sfondo" & immaginenumero
End Sub

Private Sub Cmdsucessivo_Click()
If immaginenumero < 9 Then immaginenumero = immaginenumero + 1
  Anim1.AnimatedGifPath = App.Path & "\sfondi" & "\chat" & "\immagine" & immaginenumero & ".jpg"
  Txtsfondo.Text = "sfondo" & immaginenumero
End Sub

Private Sub form_load()
  frmResize.KeepRatio = True
  frmResize.FontResize = True
  Call frmResize.InitializeResizer(Me)
  Cmdsucessivo_Click
End Sub

Private Sub Form_Resize()
  Call frmResize.FormResized(Me)
End Sub




