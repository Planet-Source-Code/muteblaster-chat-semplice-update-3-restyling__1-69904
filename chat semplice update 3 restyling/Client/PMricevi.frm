VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form PMricevi 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "ricevi singolo messaggio"
   ClientHeight    =   4815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox txtmessaggio 
      Height          =   2295
      Left            =   480
      TabIndex        =   5
      Top             =   1440
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4048
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"PMricevi.frx":0000
   End
   Begin client.chameleonButton chameleonButton1 
      Height          =   2775
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4895
      BTYPE           =   3
      TX              =   ""
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
      MPTR            =   0
      MICON           =   "PMricevi.frx":0082
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin client.CandyButton Cmdchiudi 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   975
      _ExtentX        =   1720
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
      Caption         =   "chiudi"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   3
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin client.CandyButton Cmdexit 
      Height          =   255
      Left            =   4320
      TabIndex        =   3
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ricevi messaggio"
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
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      Height          =   4815
      Left            =   0
      Top             =   0
      Width           =   5055
   End
   Begin VB.Image Image10 
      Height          =   450
      Left            =   240
      Picture         =   "PMricevi.frx":009E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4485
   End
   Begin VB.Image Image11 
      Height          =   435
      Left            =   4680
      Picture         =   "PMricevi.frx":1058
      Top             =   15
      Width           =   300
   End
   Begin VB.Image Image9 
      Height          =   450
      Left            =   0
      Picture         =   "PMricevi.frx":1766
      Top             =   0
      Width           =   345
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
      Left            =   3000
      TabIndex        =   1
      Top             =   4320
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
      Left            =   240
      TabIndex        =   0
      Top             =   4320
      Width           =   2775
   End
End
Attribute VB_Name = "PMricevi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Sub Cmdexit_Click()
 Cmdchiudi_Click
End Sub

Private Sub form_load()
 Labelora.Caption = chat.Label1.Caption
End Sub
                           
Private Sub Cmdchiudi_Click()
PMricevi.Visible = False
txtmessaggio.Text = ""
Unload PMricevi
End Sub

' questo form e' borderless ( senza bordo), impostiamo la immagine 10'
' come bordo che gli permettera' di muovere il form'
Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ReleaseCapture
SendMessage PMricevi.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub
