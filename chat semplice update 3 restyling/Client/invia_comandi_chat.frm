VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form invia_comandi_chat 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "invia "
   ClientHeight    =   5790
   ClientLeft      =   1815
   ClientTop       =   1170
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000013&
      Height          =   3495
      Left            =   1080
      ScaleHeight     =   3435
      ScaleWidth      =   2835
      TabIndex        =   14
      Top             =   5880
      Width           =   2895
   End
   Begin client.CandyButton Cmdnascondi 
      Height          =   375
      Left            =   3000
      TabIndex        =   13
      Top             =   120
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
      Caption         =   "nascondi"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   6
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin client.CandyButton cmdchiudi 
      Height          =   375
      Left            =   4200
      TabIndex        =   12
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "X"
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
   Begin VB.ListBox List1 
      Enabled         =   0   'False
      Height          =   1815
      ItemData        =   "invia_comandi_chat.frx":0000
      Left            =   120
      List            =   "invia_comandi_chat.frx":0002
      TabIndex        =   11
      Top             =   3840
      Width           =   4695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   3975
      Begin client.CandyButton Cmdprofilo 
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         Top             =   1080
         Width           =   2055
         _ExtentX        =   3625
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
         Caption         =   "guarda biglietto da visita"
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
      Begin client.CandyButton Cmdsfondo 
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
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
         Caption         =   "condividi sfondo"
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
      Begin client.CandyButton Cmdanimazioni 
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
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
         Caption         =   "condividi animazioni"
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
      Begin client.CandyButton Cmdnudge 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   120
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
         Caption         =   "nudge"
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
         Caption         =   "leggi biglietto da visita"
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Labelcondividisondo 
         BackStyle       =   0  'Transparent
         Caption         =   "condividi sfondo con un amico"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Labelanimazioni 
         BackStyle       =   0  'Transparent
         Caption         =   "invia piccole animazioni"
         Height          =   375
         Left            =   2280
         TabIndex        =   6
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Labelnudge 
         BackStyle       =   0  'Transparent
         Caption         =   "invia nudge"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   855
      End
   End
   Begin client.chameleonButton chameleonButton1 
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4695
      _ExtentX        =   8281
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
      MICON           =   "invia_comandi_chat.frx":0004
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock ws_invia_comandi_chat 
      Left            =   2400
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Shape Shape1 
      Height          =   5775
      Left            =   0
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "invia_comandi_chat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private Sub Cmdanimazioni_Click()
 SetParent animazioni_chat.hWnd, Picture1.hWnd
 animazioni_chat.Show
 Picture1.Top = 2280
 animazioni_chat.Move 0, 0
End Sub

Private Sub cmdchiudi_Click()
 Unload invia_comandi_chat
 chat.Picture17.Top = 10800
End Sub

Private Sub Cmdnascondi_Click()
 chat.Picture17.Top = 10800
End Sub

Private Sub Cmdnudge_Click()
 ws_invia_comandi_chat.SendData "(((nudge_chat)))"
End Sub

Private Sub Cmdsfondo_Click()
 sfondi.Show
 sfondi.Cmdcondividi_sfondo.Visible = True
 sfondi.Cmdanim1.Visible = False
End Sub

Private Sub Cmdprofilo_Click()
 ws_invia_comandi_chat.SendData "(((richiesta biglietto da visita)))"
End Sub

Private Sub form_load()
 ws_invia_comandi_chat.connect Trim(chat.txtIpUtente.Text), "3337" 'Connect to server
End Sub


Private Sub ws_invia_comandi_chat_DataArrival(ByVal bytesTotal As Long)
 Dim Data As String
 ws_invia_comandi_chat.GetData Data
 If Data = "(((invio biglietto)))" Then
    esegui_biglietto_da_visita.Show
 End If
    List1.AddItem Data
    esegui_biglietto_da_visita.biglietto.Text = Data
 End Sub
