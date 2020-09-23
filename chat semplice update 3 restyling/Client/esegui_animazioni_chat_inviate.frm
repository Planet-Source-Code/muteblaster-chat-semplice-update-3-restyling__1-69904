VERSION 5.00
Begin VB.Form esegui_animazioni_chat_inviate 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3045
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   3045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "animazioni inviate"
      Top             =   120
      Width           =   1695
   End
   Begin client.CandyButton Cmdexit 
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   0
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
      Caption         =   "x"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   -2147483638
      ColorButtonUp   =   -2147483638
      ColorButtonDown =   -2147483638
      BorderBrightness=   0
      ColorBright     =   -2147483638
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin client.Anim Anim1 
      Height          =   2295
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   4048
   End
   Begin client.chameleonButton chameleonButton1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   5318
      BTYPE           =   3
      TX              =   "chameleonButton1"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      MICON           =   "esegui_animazioni_chat_inviate.frx":0000
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
      Left            =   2640
      Top             =   2280
   End
   Begin VB.Shape Shape1 
      Height          =   2775
      Left            =   120
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "esegui_animazioni_chat_inviate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdexit_Click()
 Unload esegui_animazioni_chat_inviate
End Sub

Private Sub form_load()
Timer_anim1.Enabled = True
End Sub
Private Sub Timer_anim1_Timer()
Cmdexit_Click
End Sub
