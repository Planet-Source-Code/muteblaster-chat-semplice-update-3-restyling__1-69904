VERSION 5.00
Begin VB.Form sistema_di_crediti 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txtpunteggio 
      BackColor       =   &H80000003&
      Height          =   285
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Timer Timer_caricamento 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   7080
      Top             =   3960
   End
   Begin VB.Timer Timer_salvataggio 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   7800
      Top             =   3960
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000000&
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   960
      Width           =   1815
   End
   Begin client.CandyButton Cmdexit 
      Height          =   255
      Left            =   7560
      TabIndex        =   0
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "punteggio :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   4
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      Height          =   4455
      Left            =   0
      Top             =   0
      Width           =   8535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "numero di messaggi spediti"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.Image Image9 
      Height          =   450
      Left            =   0
      Picture         =   "sistema_di_crediti.frx":0000
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image11 
      Height          =   435
      Left            =   8160
      Picture         =   "sistema_di_crediti.frx":08B2
      Top             =   15
      Width           =   300
   End
   Begin VB.Image Image10 
      Height          =   450
      Left            =   240
      Picture         =   "sistema_di_crediti.frx":0FC0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7965
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
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "sistema_di_crediti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Private OldX As Integer
Private OldY As Integer

Private Sub SaveControlValues()
 Call RegSave(Text1, Text1.Text)
End Sub

Private Sub Cmdexit_Click()
 sistema_di_crediti.Visible = False
End Sub

Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 ReleaseCapture
 SendMessage sistema_di_crediti.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub

Private Sub Timer_salvataggio_Timer()
 Call SaveControlValues
 Timer_salvataggio.Enabled = False
End Sub

Private Sub Timer_caricamento_Timer()
 Text1 = RegLoad(Text1)
 Timer_caricamento.Enabled = False
End Sub
