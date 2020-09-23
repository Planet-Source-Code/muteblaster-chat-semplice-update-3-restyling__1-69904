VERSION 5.00
Begin VB.Form tempo 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "tempo"
   ClientHeight    =   2010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   2010
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin client.CandyButton cmdchiudi 
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   0
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
      Caption         =   "x"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   14737632
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16777215
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Timer Timerora 
      Interval        =   1
      Left            =   120
      Top             =   1560
   End
   Begin VB.Label Labelmioip 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Labelip 
      BackStyle       =   0  'Transparent
      Caption         =   "IP :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Labeldata 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "data :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Labelora 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Labeltimerora 
      BackStyle       =   0  'Transparent
      Caption         =   "ora :"
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
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   240
      Picture         =   "tempo.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "tempo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdchiudi_Click()
 chat.Picture9.Top = 10800
End Sub

Private Sub form_load()
 Labelmioip.Caption = login.Label1.Caption
End Sub
Private Sub Timerora_Timer()
Labelora.Caption = Time
Labeldata.Caption = Date
End Sub


