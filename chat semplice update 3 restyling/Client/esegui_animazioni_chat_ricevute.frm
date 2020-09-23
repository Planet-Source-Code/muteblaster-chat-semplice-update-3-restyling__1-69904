VERSION 5.00
Begin VB.Form esegui_animazioni_chat_ricevute 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3180
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   3180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin client.Anim Anim1 
      Height          =   2415
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   4260
   End
   Begin client.chameleonButton chameleonButton1 
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   5106
      BTYPE           =   3
      TX              =   "chameleonButton1"
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
      MICON           =   "esegui_animazioni_chat_ricevute.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer Timer_anim1 
      Interval        =   3000
      Left            =   0
      Top             =   480
   End
   Begin client.CandyButton Cmdchiudi 
      Height          =   255
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
      BackStyle       =   0  'Transparent
      Caption         =   "animazioni ricevute"
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
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      Height          =   3495
      Left            =   0
      Top             =   0
      Width           =   3135
   End
   Begin VB.Image Image9 
      Height          =   450
      Left            =   0
      Picture         =   "esegui_animazioni_chat_ricevute.frx":001C
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image11 
      Height          =   435
      Left            =   2760
      Picture         =   "esegui_animazioni_chat_ricevute.frx":08CE
      Top             =   15
      Width           =   300
   End
   Begin VB.Image Image10 
      Height          =   450
      Left            =   240
      Picture         =   "esegui_animazioni_chat_ricevute.frx":0FDC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2565
   End
End
Attribute VB_Name = "esegui_animazioni_chat_ricevute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub form_load()
 Timer_anim1.Enabled = True
 esegui_animazioni_chat_ricevute.Top = Me.Top + 2000
 esegui_animazioni_chat_ricevute.Left = Me.Left + 2000
End Sub
Private Sub Cmdchiudi_Click()
 Unload esegui_animazioni_chat_ricevute
End Sub

Private Sub Timer_anim1_Timer()
 Cmdchiudi_Click
End Sub
