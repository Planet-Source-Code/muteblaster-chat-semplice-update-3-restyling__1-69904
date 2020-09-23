VERSION 5.00
Begin VB.Form MSinvio_style 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "ms invio style"
   ClientHeight    =   6765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   10380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin client.Anim Anim1 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   11033
   End
   Begin VB.Timer Timer_movimento 
      Interval        =   1
      Left            =   360
      Top             =   4680
   End
   Begin client.CandyButton Cmdmimimizza 
      Height          =   255
      Left            =   8880
      TabIndex        =   1
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
      Caption         =   "-"
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
   Begin client.CandyButton Cmdexit 
      Height          =   255
      Left            =   9360
      TabIndex        =   2
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
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "invio messaggi privati"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin VB.Image Image9 
      Height          =   450
      Left            =   0
      Picture         =   "MSinvio_style.frx":0000
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image11 
      Height          =   435
      Left            =   9960
      Picture         =   "MSinvio_style.frx":08B2
      Top             =   15
      Width           =   300
   End
   Begin VB.Image Image10 
      Height          =   450
      Left            =   240
      Picture         =   "MSinvio_style.frx":0FC0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9765
   End
End
Attribute VB_Name = "MSinvio_style"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Sub Cmdexit_Click()
 Unload MSinvio_style
 Unload Me
End Sub

Private Sub Cmdmimimizza_Click()
 MSinvio_style.WindowState = 1
MSinvio.WindowState = 1
End Sub

Private Sub Form_Load()
 MSinvio.Show , Me
 MSinvio.Top = Me.Top + 500
 MSinvio.Left = Me.Left + 0
End Sub

' questo form e' borderless ( senza bordo), impostiamo la immagine 10'
' come bordo che gli permettera' di muovere il form'
Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage MSinvio_style.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub

Private Sub Timer_movimento_Timer()
 MSinvio.Top = Me.Top + 500
 MSinvio.Left = Me.Left + 0
 esegui_animazioni_MSinvio.Top = Me.Top + 2000
 esegui_animazioni_MSinvio.Left = Me.Left + 2500
End Sub
