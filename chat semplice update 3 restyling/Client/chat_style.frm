VERSION 5.00
Begin VB.Form chat_style 
   BorderStyle     =   0  'None
   Caption         =   "chat style"
   ClientHeight    =   11100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15315
   LinkTopic       =   "Form1"
   ScaleHeight     =   11100
   ScaleWidth      =   15315
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   10455
      Left            =   120
      ScaleHeight     =   10395
      ScaleWidth      =   14955
      TabIndex        =   2
      Top             =   480
      Width           =   15015
      Begin client.Anim Anim1 
         Height          =   10455
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   15135
         _ExtentX        =   26696
         _ExtentY        =   18441
      End
   End
   Begin client.CandyButton Criduci 
      Height          =   255
      Left            =   14040
      TabIndex        =   1
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
      Left            =   14640
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
   Begin VB.Timer Timer_movimento 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   8520
   End
   Begin VB.Shape Shape1 
      Height          =   11055
      Left            =   0
      Top             =   0
      Width           =   15255
   End
   Begin VB.Image Image9 
      Height          =   450
      Left            =   0
      Picture         =   "chat_style.frx":0000
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image10 
      Height          =   450
      Left            =   240
      Picture         =   "chat_style.frx":08B2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14685
   End
   Begin VB.Image Image11 
      Height          =   435
      Left            =   14880
      Picture         =   "chat_style.frx":186C
      Top             =   15
      Width           =   300
   End
End
Attribute VB_Name = "chat_style"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Sub Cmdexit_Click()
 chat.Timer_prepara_uscita_chat.Enabled = True
 chat.Timeresci_chat.Enabled = True
 chat.Timer_unload.Enabled = True
End Sub
' con questo comando riduciamo nella barra dei comandi i form chat'
Private Sub Criduci_Click()
 chat_style.WindowState = 1
 chat.WindowState = 1
End Sub

' all'avvio del form chat_style parte anche il form chat che viene impostato come ME'
' e viene decisa la posizione del form chat in rispetto il del form chat_style'
Private Sub Form_Load()
 'SetParent chat.hwnd, Picture1.hwnd
' chat.Show
' chat.Move 0, 0
 chat.Visible = True
 chat.Show , Me
 chat.Top = Me.Top + 500
 chat.Left = Me.Left + 50
 chat.Move 50, 500
End Sub

Private Sub Timer_movimento_Timer()
 chat.Top = Me.Top + 500
 chat.Left = Me.Left + 50
End Sub

' questo form e' borderless ( senza bordo), impostiamo la immagine 10'
' come bordo che gli permettera' di muovere il form'
Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 ReleaseCapture
 SendMessage chat_style.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub
