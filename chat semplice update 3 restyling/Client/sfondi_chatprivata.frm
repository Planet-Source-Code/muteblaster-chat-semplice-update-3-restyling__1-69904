VERSION 5.00
Begin VB.Form sfondi_chatprivata 
   BorderStyle     =   0  'None
   Caption         =   "sfondi chat privata"
   ClientHeight    =   3495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin client.CandyButton CandyButton1 
      Height          =   255
      Left            =   4320
      TabIndex        =   7
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
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.TextBox Txtsfondo 
      Height          =   285
      Left            =   2880
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin client.CandyButton Cmdprecedente 
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   3000
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
      Left            =   1800
      TabIndex        =   4
      Top             =   3000
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
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin client.CandyButton Cmdcondividi 
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "condividi sfondo"
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
   Begin client.CandyButton Cmdanim1 
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   1200
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
   Begin client.Anim Anim1 
      Height          =   2175
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   3836
   End
   Begin client.chameleonButton chameleonButton1 
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   4683
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
      MICON           =   "sfondi_chatprivata.frx":0000
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
      Height          =   3495
      Left            =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Image Image7 
      Height          =   3705
      Left            =   3000
      Picture         =   "sfondi_chatprivata.frx":001C
      Top             =   -240
      Width           =   4440
   End
   Begin VB.Image Image6 
      Height          =   3705
      Left            =   0
      Picture         =   "sfondi_chatprivata.frx":0D69
      Top             =   -240
      Width           =   4440
   End
End
Attribute VB_Name = "sfondi_chatprivata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Dim immaginenumero As Integer ' dichiariamo la variabile che identifica il numero delle immagini'

Private Sub CandyButton1_Click()
 sfondi_chatprivata.Visible = False
End Sub

Private Sub form_load()
 Cmdsucessivo_Click
End Sub

Private Sub Cmdprecedente_Click()
If immaginenumero > 1 Then immaginenumero = immaginenumero - 1 ' bisogna indicare il numero massimo di gif presenti'
                                                               ' altrimenti continuando a proseguire va' in crash'
Anim1.AnimatedGifPath = App.Path & "\sfondi" & "\privatchat" & "\immagine" & immaginenumero & ".jpg"
Txtsfondo.Text = "immagine" & immaginenumero
End Sub

Private Sub Cmdsucessivo_Click()
If immaginenumero < 5 Then immaginenumero = immaginenumero + 1
  Anim1.AnimatedGifPath = App.Path & "\sfondi" & "\privatchat" & "\immagine" & immaginenumero & ".jpg"
  Txtsfondo.Text = "sfondo" & immaginenumero
End Sub



Private Sub Cmdanim1_Click()
MSinvio_style.Anim1.AnimatedGifPath = App.Path & "\sfondi" & "\privatchat" & "\immagine" & immaginenumero & ".jpg"
MakeTransparent MSinvio.hWnd, 200
sfondi_chatprivata.Visible = False
End Sub

Private Sub Cmdcondividi_Click()
MSinvio.WsMSinvio.SendData Txtsfondo.Text
sfondi_chatprivata.Visible = False
End Sub

