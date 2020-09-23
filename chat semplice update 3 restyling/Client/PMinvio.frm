VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form PMinvio 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "invia singolo messaggio"
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin client.CandyButton Cmdsend 
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   3960
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
      Caption         =   "invia"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   2
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin RichTextLib.RichTextBox Txtsend 
      Height          =   2295
      Left            =   480
      TabIndex        =   8
      Top             =   1560
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4048
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"PMinvio.frx":0000
   End
   Begin client.chameleonButton chameleonButton1 
      Height          =   3135
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5530
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
      MPTR            =   1
      MICON           =   "PMinvio.frx":0082
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin client.CandyButton Cmdcancel 
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   4800
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
      Caption         =   "cancella"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   2
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin client.CandyButton Cmddisconnect 
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   4800
      Width           =   1095
      _ExtentX        =   1931
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
      Caption         =   "disconnetti"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   2
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   4455
      Begin VB.TextBox Txtnickamico 
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Labelmessaggio 
         BackStyle       =   0  'Transparent
         Caption         =   "stai scrivendo unmessaggio a"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2175
      End
   End
   Begin MSWinsockLib.Winsock WsPMinvio 
      Left            =   120
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin client.CandyButton Cmdexit 
      Height          =   255
      Left            =   4320
      TabIndex        =   9
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
      Caption         =   "invia messaggio"
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
      TabIndex        =   11
      Top             =   120
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      Height          =   5775
      Left            =   0
      Top             =   0
      Width           =   5055
   End
   Begin VB.Image Image10 
      Height          =   450
      Left            =   240
      Picture         =   "PMinvio.frx":009E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4485
   End
   Begin VB.Image Image11 
      Height          =   435
      Left            =   4680
      Picture         =   "PMinvio.frx":1058
      Top             =   15
      Width           =   300
   End
   Begin VB.Image Image9 
      Height          =   450
      Left            =   0
      Picture         =   "PMinvio.frx":1766
      Top             =   0
      Width           =   345
   End
   Begin VB.Label Labelinvio 
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
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label Labelora 
      BackStyle       =   0  'Transparent
      Caption         =   "messagio inviato alle :"
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
      Top             =   5400
      Width           =   2055
   End
End
Attribute VB_Name = "PMinvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Sub Cmdexit_Click()
 Cmddisconnect_Click
End Sub

Private Sub form_load()
WsPMinvio.connect Trim(chat.txtIpUtente.Text), "3333"  'connetti al server'
Cmdsend.Enabled = True 'abilita il comando send'
Txtnickamico.Text = chat.Text2.Text
End Sub
Private Sub cmdCancel_Click()
txtsend.Text = ""
End Sub

Private Sub Cmddisconnect_Click()
WsPMinvio.Close
txtsend.Enabled = True
txtsend.Text = ""
End Sub

' questo form e' borderless ( senza bordo), impostiamo la immagine 10'
' come bordo che gli permettera' di muovere il form'
Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ReleaseCapture
SendMessage PMinvio.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub

Private Sub wspminvio_connect()
 WsPMinvio.SendData " < " & chat.Txtmionick.Text & " > " & " si e' connesso per spedirti un messaggio privato " & vbCrLf & " ---------------" & vbCrLf
End Sub

Private Sub cmdsend_Click()
If WsPMinvio.State = sckConnected Then
    WsPMinvio.SendData "Client: " & txtsend.Text
End If
txtsend.Enabled = False
Labelinvio.Caption = chat.Label1.Caption
End Sub

Private Sub WsPMinvio_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ' se c'e' un problea nella connessione ci viene segnalato'
    MsgBox "non e' possibile eseguire la connessione al server....."
    Cmddisconnect.Enabled = False
End Sub
