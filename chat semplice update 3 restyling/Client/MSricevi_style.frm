VERSION 5.00
Begin VB.Form MSricevi_styleE 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   10215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer_movimento 
      Interval        =   1
      Left            =   120
      Top             =   4560
   End
   Begin client.Anim Anim1 
      Height          =   5295
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   9340
   End
   Begin client.CandyButton Cmdmimimizza 
      Height          =   255
      Left            =   9360
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ricevi messaggi privati"
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
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      Height          =   5895
      Left            =   0
      Top             =   0
      Width           =   10215
   End
   Begin VB.Image Image10 
      Height          =   450
      Left            =   240
      Picture         =   "MSricevi_style.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9645
   End
   Begin VB.Image Image11 
      Height          =   435
      Left            =   9840
      Picture         =   "MSricevi_style.frx":0FBA
      Top             =   15
      Width           =   300
   End
   Begin VB.Image Image9 
      Height          =   450
      Left            =   0
      Picture         =   "MSricevi_style.frx":16C8
      Top             =   0
      Width           =   345
   End
End
Attribute VB_Name = "MSricevi_styleE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Sub Cmdmimimizza_Click()
MSricevi_styleE.WindowState = 1
MSricevi.WindowState = 1
End Sub

Private Sub Form_Load()
 MSricevi.Show , Me
 MSricevi.Top = Me.Top + 500
 MSricevi.Left = Me.Left + 160
End Sub

' questo form e' borderless ( senza bordo), impostiamo la immagine 10'
' come bordo che gli permettera' di muovere il form'
Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage MSricevi_styleE.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub

Private Sub Timer_movimento_Timer()
MSricevi.Top = Me.Top + 500
 MSricevi.Left = Me.Left + 160
End Sub
