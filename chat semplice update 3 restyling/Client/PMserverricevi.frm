VERSION 5.00
Begin VB.Form PMserverricevi 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "RICEVUTO MESSAGGIO DAL SERVER"
   ClientHeight    =   4365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin client.CandyButton Cmdchiudi 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "chiudi"
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
   Begin VB.Timer Timer_unload_baduser 
      Interval        =   1000
      Left            =   4320
      Top             =   3720
   End
   Begin VB.TextBox Txtmessaggio 
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4695
   End
   Begin VB.Shape Shape1 
      Height          =   4335
      Left            =   0
      Top             =   0
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "RICEVI MESSAGGIO PRIVATO DAL SERVER"
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
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label Labelora 
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
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   2880
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Labelricevuto 
      BackStyle       =   0  'Transparent
      Caption         =   "hai ricevuto un messaggio alle :"
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
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   2775
   End
End
Attribute VB_Name = "PMserverricevi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub form_load()
 Labelora.Caption = chat.Label1.Caption
End Sub

Private Sub cmdchiudi_Click()
PMricevi.Visible = False
Txtmessaggio.Text = ""
Unload PMricevi
End Sub

' se il server invia un comando di penalizzazione, tipo ban o chiusura dle programma'
' un timer si preoccupera' di chiudere il form di modo che non venga lasciato aperto volutamente'
Private Sub Timer_unload_baduser_Timer()
Unload PMserverricevi
End Sub
