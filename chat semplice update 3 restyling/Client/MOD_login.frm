VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form MOD_login 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "login"
   ClientHeight    =   3405
   ClientLeft      =   4725
   ClientTop       =   3180
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000B&
      Height          =   2175
      Left            =   600
      TabIndex        =   13
      Top             =   6000
      Width           =   4935
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000A&
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
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "la password e' sbagliata"
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   465
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "errore"
         Top             =   360
         Width           =   1215
      End
      Begin client.chameleonButton chameleonButton2 
         Height          =   2175
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   3836
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
         MICON           =   "MOD_login.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   3000
         Left            =   4440
         Top             =   240
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Height          =   2175
      Left            =   600
      TabIndex        =   7
      Top             =   3720
      Width           =   4935
      Begin client.CandyButton Cmdverifica 
         Height          =   375
         Left            =   2640
         TabIndex        =   12
         Top             =   1320
         Width           =   855
         _ExtentX        =   1508
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
         Caption         =   "verifica"
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
      Begin VB.TextBox Txtpassword 
         BackColor       =   &H80000013&
         Height          =   375
         Left            =   1320
         TabIndex        =   10
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000A&
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "scrivi la password di sicurezza"
         Top             =   120
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Visible         =   0   'False
         Width           =   2175
      End
      Begin client.chameleonButton chameleonButton1 
         Height          =   2175
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   3836
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
         MICON           =   "MOD_login.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin client.CandyButton Cmdconnect 
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
      _ExtentX        =   2143
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
      Caption         =   "connetti"
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
   Begin VB.Frame framelogin 
      BackColor       =   &H80000013&
      Caption         =   "login"
      Height          =   2055
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   4935
      Begin VB.TextBox Txtfrase 
         Height          =   495
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox Txtnick 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Labelfrase 
         BackStyle       =   0  'Transparent
         Caption         =   "frase tipica"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Labelnick 
         BackStyle       =   0  'Transparent
         Caption         =   "MODERATORE"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSWinsockLib.Winsock WS 
      Left            =   120
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin client.CandyButton Cmdexit 
      Height          =   255
      Left            =   5520
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
      Height          =   3375
      Left            =   0
      Top             =   0
      Width           =   6255
   End
   Begin VB.Image Image9 
      Height          =   450
      Left            =   0
      Picture         =   "MOD_login.frx":0038
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image11 
      Height          =   435
      Left            =   5880
      Picture         =   "MOD_login.frx":08EA
      Top             =   15
      Width           =   300
   End
   Begin VB.Image Image10 
      Height          =   450
      Left            =   240
      Picture         =   "MOD_login.frx":0FF8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5685
   End
End
Attribute VB_Name = "MOD_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 Txtnick.Text = login.Txtnick.Text
 If Txtpassword.Text = "" Then
    Cmdverifica.Enabled = False
 End If
 Frame1.Top = 1080
 Frame1.Left = 600
End Sub

Private Sub Cmdexit_Click()
 Unload MOD_login
End Sub

Private Sub Cmdverifica_Click()
 If Txtpassword.Text = Text2.Text Then
    Frame1.Visible = False
 Else
    Frame2.Left = 600
    Frame2.Top = 1080
    Timer1.Enabled = True
 End If
End Sub

Private Sub Cmdconnect_Click()
 WS.connect login.txtIP.Text, 4000 ' Connects to the server
End Sub

Private Sub Timer1_Timer()
 Frame2.Left = 360
 Frame2.Top = 6000
 Timer1.Enabled = False
End Sub

Private Sub Txtpassword_Change()
 If Txtpassword.Text = "" Then
    Cmdverifica.Enabled = False
 Else
    Cmdverifica.Enabled = True
 End If
End Sub

' informazioni dopo la connessione e prima di accedere'
' alla chat'
Private Sub WS_connect()
MsgBox "sei connesso a : " & WS.RemoteHost, vbInformation, "connesso"
        MOD_login.WS.SendData "@CONNECT:" & Txtnick.Text
MOD_login.Visible = False
MOD_chat.Show
End Sub

Private Sub WS_DataArrival(ByVal bytesTotal As Long)
Dim txt As String
Dim Splittati() As String
WS.GetData txt, vbString '
If Mid(txt, 1, 6) = "@LISTC" Then
    MOD_chat.listusers.Clear
    Splittati() = Split(Mid(txt, 7), "@")
    For I = LBound(Splittati) + 1 To UBound(Splittati)
        Me.Caption = Splittati(I)
            MOD_chat.listusers.AddItem Mid(Splittati(I), InStr(1, Splittati(I), "LIST:") + 5)
    Next I
    Exit Sub
End If

MOD_chat.txtchat.Text = MOD_chat.txtchat.Text + txt + vbCrLf ' il txtcgat contiene sia i messaggi inviati che ricevuti'
End Sub
