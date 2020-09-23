VERSION 5.00
Begin VB.Form cambianick 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "cambia il nick"
   ClientHeight    =   2340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin client.chameleonButton Cmdcambianick 
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "cambio nick"
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
      MICON           =   "cambianick.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Framecambionick 
      BackColor       =   &H80000013&
      Caption         =   "cambionick"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4455
      Begin VB.TextBox Txtnuovonick 
         Height          =   285
         Left            =   2280
         TabIndex        =   2
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox Txtnickattuale 
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Labelnuovonick 
         BackStyle       =   0  'Transparent
         Caption         =   "nuovo nick"
         Height          =   255
         Left            =   2520
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Labelnickattuale 
         BackStyle       =   0  'Transparent
         Caption         =   "nick attuale"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin client.CandyButton Cmdexit 
      Height          =   255
      Left            =   4080
      TabIndex        =   5
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
      Height          =   2295
      Left            =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Image Image10 
      Height          =   450
      Left            =   240
      Picture         =   "cambianick.frx":001C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4125
   End
   Begin VB.Image Image11 
      Height          =   435
      Left            =   4320
      Picture         =   "cambianick.frx":0FD6
      Top             =   15
      Width           =   300
   End
   Begin VB.Image Image9 
      Height          =   450
      Left            =   0
      Picture         =   "cambianick.frx":16E4
      Top             =   0
      Width           =   345
   End
End
Attribute VB_Name = "cambianick"
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
 chat.Picture19.Top = 10800
End Sub

Private Sub form_load()
 Txtnickattuale.Text = chat.Txtmionick.Text
End Sub

Private Sub Cmdcambianick_Click()
On Error Resume Next
chat.Txtmionick.Text = Txtnuovonick.Text
login.WS.SendData Chr(127) & "nick:" & login.Txtnick.Text & Chr(127) & "cambionick:" & " <<ha cambiato nick in: " & Txtnuovonick.Text
cambianick.Visible = False
End Sub


Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ReleaseCapture
SendMessage cambianick.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub
