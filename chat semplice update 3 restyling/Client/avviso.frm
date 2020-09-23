VERSION 5.00
Begin VB.Form avviso 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "messaggio di avviso"
   ClientHeight    =   1725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin client.CandyButton Cmdok 
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
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
   Begin VB.Timer Timer_unload 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   360
      Top             =   1200
   End
   Begin client.CandyButton Cmdexit 
      Height          =   255
      Left            =   5280
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
   Begin VB.Shape Shape1 
      Height          =   1695
      Left            =   0
      Top             =   0
      Width           =   5895
   End
   Begin VB.Image Image9 
      Height          =   450
      Left            =   0
      Picture         =   "avviso.frx":0000
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image11 
      Height          =   435
      Left            =   5520
      Picture         =   "avviso.frx":08B2
      Top             =   15
      Width           =   300
   End
   Begin VB.Image Image10 
      Height          =   450
      Left            =   240
      Picture         =   "avviso.frx":0FC0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5325
   End
   Begin VB.Label Labelmessaggio 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label Labelavviso 
      BackStyle       =   0  'Transparent
      Caption         =   "avviso :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "avviso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
cmdOk_Click
End Sub

Private Sub cmdOk_Click()
Unload avviso
End Sub

Private Sub Timer_unload_Timer()
 cmdOk_Click
End Sub

 Private Sub Form_Load()
 Timer_unload.Enabled = True
 End Sub
