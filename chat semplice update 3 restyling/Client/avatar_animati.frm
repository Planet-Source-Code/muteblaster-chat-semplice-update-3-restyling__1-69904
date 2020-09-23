VERSION 5.00
Begin VB.Form avatar_animati 
   BackColor       =   &H8000000D&
   Caption         =   "scegli avatar animati"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   3390
   StartUpPosition =   3  'Windows Default
   Begin client.CandyButton Cmdprecedente 
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   2640
      Width           =   735
      _ExtentX        =   1296
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
      Caption         =   "<<<<"
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
   Begin VB.TextBox Txtavatar 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   1560
      Width           =   615
   End
   Begin client.CandyButton Cmdsucessivo 
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   2640
      Width           =   735
      _ExtentX        =   1296
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
      Caption         =   ">>>>"
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
      Height          =   1695
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2990
   End
   Begin client.CandyButton CandyButton1 
      Height          =   2175
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   3836
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "CandyButton1"
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
End
Attribute VB_Name = "avatar_animati"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim immaginenumero As Integer ' dichiariamo la variabile che identifica il numero delle immagini'

Private Sub Cmdprecedente_Click()
If immaginenumero > 1 Then immaginenumero = immaginenumero - 1 ' bisogna indicare il numero massimo di gif presenti'
                                                               ' altrimenti continuando a proseguire va' in crash'
Anim1.AnimatedGifPath = App.Path & "\avatar" & "\immagine" & immaginenumero & ".gif"
Txtavatar.Text = immaginenumero
End Sub

Private Sub Cmdsucessivo_Click()
If immaginenumero < 79 Then immaginenumero = immaginenumero + 1
  Anim1.AnimatedGifPath = App.Path & "\avatar" & "\immagine" & immaginenumero & ".gif"
  Txtavatar.Text = immaginenumero
End Sub

 
 
