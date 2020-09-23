VERSION 5.00
Begin VB.Form informazioni_chiusura 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5910
   ClientLeft      =   9090
   ClientTop       =   1170
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List6 
      BackColor       =   &H8000000A&
      Enabled         =   0   'False
      Height          =   2205
      ItemData        =   "informazioni_chiusura.frx":0000
      Left            =   3960
      List            =   "informazioni_chiusura.frx":000D
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ListBox List5 
      BackColor       =   &H8000000A&
      Enabled         =   0   'False
      Height          =   2595
      ItemData        =   "informazioni_chiusura.frx":006F
      Left            =   2040
      List            =   "informazioni_chiusura.frx":0094
      TabIndex        =   4
      Top             =   3240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ListBox List4 
      BackColor       =   &H8000000A&
      Enabled         =   0   'False
      Height          =   2595
      ItemData        =   "informazioni_chiusura.frx":018E
      Left            =   120
      List            =   "informazioni_chiusura.frx":01C5
      TabIndex        =   3
      Top             =   3240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ListBox List3 
      BackColor       =   &H8000000A&
      Enabled         =   0   'False
      Height          =   2595
      ItemData        =   "informazioni_chiusura.frx":0330
      Left            =   3960
      List            =   "informazioni_chiusura.frx":036D
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ListBox List2 
      BackColor       =   &H8000000A&
      Enabled         =   0   'False
      Height          =   2595
      ItemData        =   "informazioni_chiusura.frx":04FB
      Left            =   2040
      List            =   "informazioni_chiusura.frx":052F
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ListBox List1 
      BackColor       =   &H8000000A&
      Enabled         =   0   'False
      Height          =   2595
      ItemData        =   "informazioni_chiusura.frx":06EF
      Left            =   120
      List            =   "informazioni_chiusura.frx":072C
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin client.CandyButton Cmdexit 
      Height          =   255
      Left            =   5160
      TabIndex        =   6
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
      Caption         =   "X"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
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
      Caption         =   "informazioni sulla chiusura del programma"
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
      Left            =   360
      TabIndex        =   7
      Top             =   120
      Width           =   4575
   End
   Begin VB.Shape Shape1 
      Height          =   5895
      Left            =   0
      Top             =   0
      Width           =   5895
   End
   Begin VB.Image Image10 
      Height          =   450
      Left            =   240
      Picture         =   "informazioni_chiusura.frx":08EA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5325
   End
   Begin VB.Image Image11 
      Height          =   435
      Left            =   5520
      Picture         =   "informazioni_chiusura.frx":18A4
      Top             =   15
      Width           =   300
   End
   Begin VB.Image Image9 
      Height          =   450
      Left            =   0
      Picture         =   "informazioni_chiusura.frx":1FB2
      Top             =   0
      Width           =   345
   End
End
Attribute VB_Name = "informazioni_chiusura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "USER32" () As Long
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Private OldX As Integer
Private OldY As Integer

Private Sub form_load()
 Call AddHScroll(List1)
 Call AddHScroll(List2)
 Call AddHScroll(List3)
 Call AddHScroll(List4)
 Call AddHScroll(List5)
 Call AddHScroll(List6)
End Sub


Private Sub Cmdexit_Click()
 informazioni_chiusura.Visible = False
End Sub

Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 ReleaseCapture
 SendMessage informazioni_chiusura.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 ReleaseCapture
 SendMessage informazioni_chiusura.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub
