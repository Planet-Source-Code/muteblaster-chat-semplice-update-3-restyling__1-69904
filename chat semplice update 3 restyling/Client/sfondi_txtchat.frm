VERSION 5.00
Begin VB.Form sfondi_txtchat 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin client.CandyButton CandyButton1 
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   0
      Width           =   495
      _ExtentX        =   873
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
   Begin client.CandyButton Cmdprecedente 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3120
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
   Begin client.CandyButton Cmdanim1 
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   1560
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
   Begin VB.TextBox Txtsfondo 
      Height          =   285
      Left            =   3600
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin client.CandyButton Cmdsucessivo 
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   3120
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
   Begin client.Anim Anim1 
      Height          =   2295
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   4048
   End
   Begin client.chameleonButton chameleonButton1 
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   4895
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
      MICON           =   "sfondi_txtchat.frx":0000
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
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   3240
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
      Caption         =   "chiudi"
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
      Height          =   3735
      Left            =   0
      Top             =   0
      Width           =   4575
   End
   Begin VB.Image Image14 
      Height          =   3705
      Left            =   1680
      Picture         =   "sfondi_txtchat.frx":001C
      Top             =   0
      Width           =   4440
   End
   Begin VB.Image Image13 
      Height          =   3705
      Left            =   0
      Picture         =   "sfondi_txtchat.frx":0D69
      Top             =   0
      Width           =   4440
   End
End
Attribute VB_Name = "sfondi_txtchat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim immaginenumero As Integer ' dichiariamo la variabile che identifica il numero delle immagini'

Private Sub CandyButton1_Click()
  chat.Picture16.Top = 10800
End Sub

Private Sub Cmdanim1_Click()
chat.Anim1.AnimatedGifPath = App.Path & "\sfondi" & "\txtchat" & "\immagine" & immaginenumero & ".jpg"
sfondi_txtchat.Visible = False
End Sub

Private Sub Cmdchiudi_Click()
 sfondi_txtchat.Visible = False
End Sub

Private Sub Cmdprecedente_Click()
If immaginenumero > 1 Then immaginenumero = immaginenumero - 1 ' bisogna indicare il numero massimo di gif presenti'
                                                               ' altrimenti continuando a proseguire va' in crash'
Anim1.AnimatedGifPath = App.Path & "\sfondi" & "\txtchat" & "\immagine" & immaginenumero & ".jpg"
Txtsfondo.Text = "immagine" & immaginenumero
End Sub

Private Sub Cmdsucessivo_Click()
If immaginenumero < 5 Then immaginenumero = immaginenumero + 1
  Anim1.AnimatedGifPath = App.Path & "\sfondi" & "\txtchat" & "\immagine" & immaginenumero & ".jpg"
  Txtsfondo.Text = "sfondo" & immaginenumero
End Sub

Private Sub form_load()
 Cmdsucessivo_Click
End Sub

