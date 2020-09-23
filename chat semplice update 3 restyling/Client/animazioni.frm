VERSION 5.00
Begin VB.Form animazioni_MSinvio 
   BorderStyle     =   0  'None
   Caption         =   "animazioni"
   ClientHeight    =   3645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin client.CandyButton Cmdanim1 
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   1200
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
   Begin VB.TextBox Txtanimazione 
      Height          =   285
      Left            =   3240
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin client.CandyButton Cmdprecedente 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   3120
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
      Left            =   2160
      TabIndex        =   2
      Top             =   3120
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
      Height          =   1935
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   3413
   End
   Begin client.chameleonButton chameleonButton1 
      Height          =   2415
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   4260
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
      MICON           =   "animazioni.frx":0000
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
      Height          =   3615
      Left            =   0
      Top             =   0
      Width           =   4335
   End
   Begin VB.Image Image9 
      Height          =   450
      Left            =   0
      Picture         =   "animazioni.frx":001C
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image11 
      Height          =   435
      Left            =   3960
      Picture         =   "animazioni.frx":08CE
      Top             =   15
      Width           =   300
   End
   Begin VB.Image Image10 
      Height          =   450
      Left            =   240
      Picture         =   "animazioni.frx":0FDC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3765
   End
   Begin VB.Image Image7 
      Height          =   3705
      Left            =   3840
      Picture         =   "animazioni.frx":1F96
      Top             =   0
      Width           =   4440
   End
   Begin VB.Image Image6 
      Height          =   3705
      Left            =   0
      Picture         =   "animazioni.frx":2CE3
      Top             =   0
      Width           =   4440
   End
End
Attribute VB_Name = "animazioni_MSinvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' questo form serve per le animazioni da spedire nella messaggistica privata'
' le animazioni sono realizzate con un controllo utente apposito......'
' ogni controllo utente e' fornita di un comand button ok che da' l'avvio a tutti '
' i processi necessari per il corretto svcolgimento della applicazione'
' infatti la animazione nel form msinvio e' impostata con visibilita' falsa'
' il comando ok la riporta visibile e fa' partir il timer che ne regola la durata'
' il timer e' impostato a 3 secondi, perche' ritenevo fosse una durata opportuna'
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Private OldX As Integer
Private OldY As Integer

Dim immaginenumero As Integer ' dichiariamo la variabile che identifica il numero delle immagini'

Private Sub Cmdanim1_Click()
esegui_animazioni_MSinvio.Show
esegui_animazioni_MSinvio.Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\MS" & "\immagine" & immaginenumero & ".gif"
MSinvio.WsMSinvio.SendData Txtanimazione.Text
End Sub

Private Sub Cmdprecedente_Click()
If immaginenumero > 1 Then immaginenumero = immaginenumero - 1 ' bisogna indicare il numero massimo di gif presenti'
                                                               ' altrimenti continuando a proseguire va' in crash'
Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\chat" & "\immagine" & immaginenumero & ".gif"
Txtanimazione.Text = "animazione" & immaginenumero
End Sub

Private Sub Cmdsucessivo_Click()
If immaginenumero < 6 Then immaginenumero = immaginenumero + 1
  Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\MS" & "\immagine" & immaginenumero & ".gif"
  Txtanimazione.Text = "animazione" & immaginenumero
End Sub

Private Sub Form_Load()
 Cmdsucessivo_Click
End Sub

Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ReleaseCapture
SendMessage animazioni_MSinvio.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub
