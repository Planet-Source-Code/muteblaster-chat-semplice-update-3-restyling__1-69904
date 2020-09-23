VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form MODERAZIONE 
   BackColor       =   &H80000013&
   Caption         =   "MODERAZIONE"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   6450
   StartUpPosition =   3  'Windows Default
   Begin client.CandyButton Cmdsend 
      Height          =   255
      Left            =   3000
      TabIndex        =   13
      Top             =   3480
      Width           =   735
      _ExtentX        =   1296
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
      Caption         =   "invia"
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
   Begin client.CandyButton Cmdcancel 
      Height          =   255
      Left            =   1920
      TabIndex        =   12
      Top             =   3480
      Width           =   975
      _ExtentX        =   1720
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
      Caption         =   "cancella"
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
   Begin client.CandyButton Cmddisconnect 
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   3480
      Width           =   1095
      _ExtentX        =   1931
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
      Caption         =   "disconnetti"
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
   Begin client.Anim Anim2 
      Height          =   375
      Left            =   4320
      TabIndex        =   10
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
   End
   Begin client.Anim Anim1 
      Height          =   1695
      Left            =   4320
      TabIndex        =   9
      Top             =   480
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   2990
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3615
      Begin VB.TextBox Txtnickamico 
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Labelmessaggio 
         BackStyle       =   0  'Transparent
         Caption         =   "stai scrivendo unmessaggio a"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame_comandi 
      BackColor       =   &H80000013&
      Caption         =   "COMANDI MODERAZIONE"
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   4320
      Width           =   4095
      Begin client.CandyButton Cmdbloccachat 
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "blocca chat"
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
      Begin client.CandyButton Cmdban 
         Height          =   495
         Left            =   1080
         TabIndex        =   15
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "bannaggio"
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
      Begin client.CandyButton Cmdend 
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "end"
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
      Begin VB.CommandButton Cmdblocca_privat 
         Caption         =   "blocca attivita' private"
         Height          =   495
         Left            =   2640
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
   End
   Begin RichTextLib.RichTextBox Txtsend 
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   2990
      _Version        =   393217
      TextRTF         =   $"MODERAZIONE.frx":0000
   End
   Begin MSWinsockLib.Winsock WsMODERATORE 
      Left            =   120
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
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
      Left            =   2520
      TabIndex        =   8
      Top             =   3960
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
      TabIndex        =   7
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "INVIA UN MESSAGGIO ALL'UTENTE  DA AVVISARE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "MODERAZIONE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdban_Click()
WsMODERATORE.SendData "(((ban)))"
End Sub

Private Sub Cmdblocca_privat_Click()
WsMODERATORE.SendData "(((block privat)))"
End Sub

Private Sub Cmdbloccachat_Click()
WsMODERATORE.SendData "(((lock chat)))"
End Sub

Private Sub Cmdend_Click()
WsMODERATORE.SendData "(((end)))"
End Sub

Private Sub Form_Load()
WsMODERATORE.connect Trim(chat.txtIpUtente.Text), "3336"  'connetti al server'
Cmdsend.Enabled = True 'abilita il comando send'
Txtnickamico.Text = " ADMIN "
Anim1.AnimatedGifPath = App.Path & "\immagini varie" & "\moderatore.jpg"
Anim2.AnimatedGifPath = App.Path & "\immagini varie" & "\moderatore_scritta.gif"
End Sub
Private Sub cmdCancel_Click()
Txtsend.Text = ""
End Sub

Private Sub Cmddisconnect_Click()
WsMODERATORE.Close
Txtsend.Enabled = True
Txtsend.Text = ""
End Sub
Private Sub WsMODERATORE_connect()
 avviso.Labelmessaggio.Caption = "puoi spedire il messaggio al server"
End Sub

Private Sub cmdSend_Click()
If WsMODERATORE.State = sckConnected Then
    WsMODERATORE.SendData " ADMIN   " & " > " & Txtsend.Text
End If
Txtsend.Enabled = False
Labelinvio.Caption = server.Label1.Caption
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 Cmddisconnect_Click
End Sub

Private Sub WsMODERATORE_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ' se c'e' un problea nella connessione ci viene segnalato'
    MsgBox "non e' possibile eseguire la connessione al server....."
    Cmddisconnect.Enabled = False
End Sub



