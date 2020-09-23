VERSION 5.00
Begin VB.Form blocca_sblocca 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "blocca e sblocca "
   ClientHeight    =   2820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   ScaleHeight     =   2820
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin client.CandyButton CmdOK 
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   2400
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
   Begin client.CandyButton Cmdtxtsendsblocca 
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "sblocca"
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
   Begin client.CandyButton Cmdtxtsendblocca 
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "blocca"
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
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   240
      Top             =   2160
   End
   Begin client.CandyButton Cmdexit 
      Height          =   255
      Left            =   3600
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
   Begin VB.Shape Shape1 
      Height          =   2775
      Left            =   0
      Top             =   0
      Width           =   4335
   End
   Begin VB.Image Image10 
      Height          =   450
      Left            =   240
      Picture         =   "blocca_sblocca.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3765
   End
   Begin VB.Image Image11 
      Height          =   435
      Left            =   3960
      Picture         =   "blocca_sblocca.frx":0FBA
      Top             =   15
      Width           =   300
   End
   Begin VB.Image Image9 
      Height          =   450
      Left            =   0
      Picture         =   "blocca_sblocca.frx":16C8
      Top             =   0
      Width           =   345
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "se voletche che qualcuno non invii messaggi in vostra assenza bloccate la possibilita' di scrivere messaggi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   3495
   End
End
Attribute VB_Name = "blocca_sblocca"
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

Private Sub Cmdexit_Click()
 Cmdok_Click
End Sub

Private Sub Cmdok_Click()
 chat.Picture10.Top = 10800
End Sub

' ho fatto questo semplice form per evitare che vengano spediti messaggi accidentalmente'
' un po' mi sono ispirato al blocco tastiera del cellulare'
Private Sub Cmdtxtsendblocca_Click()
On Error Resume Next
chat.txtsend.Locked = True
End Sub

 ' riabilitiamo la scrittura dei messaggi'
Private Sub Cmdtxtsendsblocca_Click()
On Error Resume Next
chat.txtsend.Locked = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Cancel = 1
 Cmdok_Click
End Sub

Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ReleaseCapture
SendMessage blocca_sblocca.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub
Private Sub Timer1_Timer()
If blocca_sblocca.Visible = True Then
 Cmdok_Click
End If
End Sub
