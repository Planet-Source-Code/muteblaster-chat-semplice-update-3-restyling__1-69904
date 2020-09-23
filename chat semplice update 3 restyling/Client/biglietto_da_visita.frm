VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form biglietto_da_visita 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   ScaleHeight     =   7965
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin client.CandyButton Cmdsalva 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   7200
      Width           =   1215
      _ExtentX        =   2143
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
      Caption         =   "salva"
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
   Begin RichTextLib.RichTextBox biglietto 
      Height          =   4935
      Left            =   720
      TabIndex        =   1
      Top             =   1680
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   8705
      _Version        =   393217
      TextRTF         =   $"biglietto_da_visita.frx":0000
   End
   Begin client.chameleonButton chameleonButton1 
      Height          =   5655
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   9975
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
      MICON           =   "biglietto_da_visita.frx":0082
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin client.CandyButton Cmdexit 
      Height          =   255
      Left            =   4200
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "scrivi il tuo biglietto da visita, da mettere a disposzione degli utenti"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   4095
   End
   Begin VB.Shape Shape1 
      Height          =   7935
      Left            =   0
      Top             =   0
      Width           =   4815
   End
   Begin VB.Image Image10 
      Height          =   450
      Left            =   240
      Picture         =   "biglietto_da_visita.frx":009E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4245
   End
   Begin VB.Image Image11 
      Height          =   435
      Left            =   4440
      Picture         =   "biglietto_da_visita.frx":1058
      Top             =   15
      Width           =   300
   End
   Begin VB.Image Image9 
      Height          =   450
      Left            =   0
      Picture         =   "biglietto_da_visita.frx":1766
      Top             =   0
      Width           =   345
   End
End
Attribute VB_Name = "biglietto_da_visita"
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

Private Sub cmdExit_Click()
 biglietto_da_visita.Visible = False
End Sub

Private Sub CMDSALVA_Click()
 If biglietto.Text = "" Then
    opzioni.Check3 = 0
    opzioni.Check3.Enabled = False
 Else
    opzioni.Check3.Enabled = True
    opzioni.Check3 = 1
 End If
 SaveText biglietto, App.Path & "\informazioni utente\biglietto da visita\biglietto.txt"
End Sub

Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ReleaseCapture
SendMessage biglietto_da_visita.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub
