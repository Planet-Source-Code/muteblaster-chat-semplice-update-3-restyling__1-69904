VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form MSinvio 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "chat privata invio"
   ClientHeight    =   6525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin client.CandyButton Cmddisconnect 
      Height          =   495
      Left            =   5520
      TabIndex        =   17
      Top             =   1200
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "chiudi conversazione"
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
   Begin client.CandyButton Cmdsend 
      Height          =   375
      Left            =   6720
      TabIndex        =   16
      Top             =   5880
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
      Caption         =   "send"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Height          =   4935
      Left            =   7560
      TabIndex        =   13
      Top             =   1440
      Width           =   2655
      Begin VB.PictureBox Picavatar 
         BackColor       =   &H80000013&
         Height          =   1575
         Left            =   480
         ScaleHeight     =   1515
         ScaleWidth      =   1515
         TabIndex        =   15
         Top             =   360
         Width           =   1575
      End
      Begin VB.PictureBox Picmioavatar 
         BackColor       =   &H80000013&
         Height          =   1575
         Left            =   480
         ScaleHeight     =   1515
         ScaleWidth      =   1515
         TabIndex        =   14
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Shape Shape1 
         Height          =   1815
         Left            =   360
         Shape           =   5  'Rounded Square
         Top             =   240
         Width           =   1815
      End
      Begin VB.Image shape5 
         Height          =   615
         Left            =   2160
         Picture         =   "MSinvio.frx":0000
         Top             =   360
         Width           =   255
      End
      Begin VB.Shape Shape2 
         Height          =   1815
         Left            =   360
         Shape           =   5  'Rounded Square
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Image shap6 
         Height          =   615
         Left            =   2160
         Picture         =   "MSinvio.frx":0896
         Top             =   2880
         Width           =   255
      End
   End
   Begin client.CandyButton Cmdnudge 
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   5160
      Width           =   735
      _ExtentX        =   1296
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
   Begin client.CandyButton Cmdanimazioni 
      Height          =   375
      Left            =   1200
      TabIndex        =   11
      Top             =   5160
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
      Caption         =   "anmazioni"
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
   Begin VB.Timer Timer_movimento 
      Interval        =   1
      Left            =   6840
      Top             =   5160
   End
   Begin client.CandyButton Cmdsfondi 
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   5160
      Width           =   735
      _ExtentX        =   1296
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
      Caption         =   "sfondi"
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
   Begin client.CandyButton Cmdwebcam 
      Height          =   375
      Left            =   5640
      TabIndex        =   9
      Top             =   5160
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
      Caption         =   "webcam"
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
   Begin VB.Frame Framecontatti 
      BackColor       =   &H80000013&
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   5175
      Begin VB.TextBox Txtmionick 
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Labelamico 
         BackStyle       =   0  'Transparent
         Caption         =   "amico"
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Labelmionick 
         BackStyle       =   0  'Transparent
         Caption         =   "io"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   135
      End
   End
   Begin VB.TextBox Txtsend 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   5640
      Width           =   6375
   End
   Begin VB.ListBox Listms 
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   7215
   End
   Begin MSWinsockLib.Winsock WsMSinvio 
      Left            =   5400
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image Image7 
      Height          =   315
      Left            =   10080
      Picture         =   "MSinvio.frx":112C
      Top             =   0
      Width           =   180
   End
   Begin VB.Image Image4 
      Height          =   315
      Left            =   7080
      Picture         =   "MSinvio.frx":1462
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3000
   End
   Begin VB.Image Image3 
      Height          =   870
      Left            =   0
      Picture         =   "MSinvio.frx":1894
      Top             =   0
      Width           =   7155
   End
   Begin VB.Shape Shape3 
      Height          =   6495
      Left            =   0
      Top             =   0
      Width           =   10335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ora:"
      Height          =   255
      Left            =   7440
      TabIndex        =   8
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Labelora 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   8160
      TabIndex        =   7
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "MSinvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Dim frmResize As New ControlResizer
Dim Newalert As New cpopup ' dichiariamo la variabile per aprire i popup '
                           ' richiamandoli dal modulo cpopup'

Private Sub Anim1_GotFocus()

End Sub

Private Sub Cmdanimazioni_Click()
animazioni_MSinvio.Show
End Sub

Private Sub Cmdmimimizza_Click()

End Sub

Private Sub Cmdnudge_Click()
nudge.Timernudge1.Enabled = True
Txtsend.Text = "(((($$$$ nudge ££££)))))"  'identifichaimo il nudge con una espessione impossibile '
                                           ' da ripetere casualmente'

WsMSinvio.SendData "(((($$$$ nudge ££££)))))"   ' dopo di che inviamo in modo tale che dall'altra parte ricevendo questo insieme di simboli'
                                                ' avviino il nudge'
End Sub

Private Sub Cmdsfondi_Click()
 sfondi_chatprivata.Show , Me
 sfondi_chatprivata.Top = Me.Top + 2200
 sfondi_chatprivata.Left = Me.Left + 2300
End Sub

Private Sub form_load()
Txtmionick.Text = chat.Txtmionick.Text ' all'avvio del form metto a video il mio nick prendendolo dal login'
Text2.Text = chat.Text2.Text
Labelora.Caption = chat.Label1.Caption
Picavatar.Picture = LoadPicture(App.Path & "\avatar" & "\immagine" & informazioniutente.Txtavataramico & ".gif")
Picmioavatar.Picture = LoadPicture(App.Path & "\avatar" & "\immagine" & avatar.Txtavatar & ".gif")
WsMSinvio.connect Trim(chat.txtIpUtente.Text), "3334" 'connetti al server'
Cmdsend.Enabled = True 'abilita il comando send'
 
  frmResize.KeepRatio = True
  frmResize.FontResize = True
  Call frmResize.InitializeResizer(Me)
End Sub

Private Sub Form_Resize()

  Call frmResize.FormResized(Me)
    
End Sub

Private Sub Cmddisconnect_Click()
Listms.AddItem " la conversazione tra > " & Txtmionick.Text & "   e   " & Text2.Text & " e' terminata alle " & Labelora.Caption & "< del>" & chat.Label2.Caption
WsMSinvio.Close ' chiudiamo il winsock'
End Sub

Private Sub cmdsend_Click()
If WsMSinvio.State = sckConnected Then ' verifica se sei connesso'
    If Txtsend.Text = "" Then 'verifica se hai scritto qualcosa nel txtsend'
        Else
            WsMSinvio.SendData Txtmionick.Text & " > " & Txtsend.Text 'invia il testo'
            Listms.AddItem Txtmionick.Text & " > " & Txtsend.Text 'metti a video il messaggio'
            Txtsend.Text = "" 'cancella in txtsend'
    End If
End If
End Sub

' questo form e' borderless ( senza bordo), impostiamo la immagine 10'
' come bordo che gli permettera' di muovere il form'
Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ReleaseCapture
SendMessage MSinvio_style.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
SendMessage MSinvio.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub

Private Sub shap6_Click()
If Picmioavatar.Visible = True Then
   Picmioavatar.Visible = False
 Else
   Picmioavatar.Visible = True
 End If
End Sub

Private Sub shape5_Click()
If Picavatar.Visible = True Then
   Picavatar.Visible = False
 Else
   Picavatar.Visible = True
 End If
End Sub

Private Sub Timer_movimento_Timer()
 sfondi_chatprivata.Top = Me.Top + 2100
 sfondi_chatprivata.Left = Me.Left + 2300
End Sub

' se la conessione va' a buon fine vieni avvisato da un messaggio'
' poi viene inviato un messaggio con il mio nick '
Private Sub WsMSinvio_connect()
 WsMSinvio.SendData " SI E' CONNESSO " & " > " & Txtmionick.Text
 Listms.AddItem " inizio chat privata >  alle  " & Labelora.Caption & "< del>" & chat.Label2.Caption & "< tra >" & Txtmionick.Text & "   e   " & Text2.Text & " --------------------"
 Listms.AddItem " --------------------"
End Sub

Private Sub WsMSinvio_close()
 MsgBox " la connessione con " & " : " & Text2.Text & "  " & " e' caduta "
 login.WS.SendData Txtmionick.Text & " << ha concluso la chat privata con <<: " & Text2.Text
  Cmddisconnect_Click
End Sub

Private Sub WsMSinvio_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
WsMSinvio.GetData Data
If Data = "animazione1" Then
 esegui_animazioni_MSinvio.Show
 esegui_animazioni_MSinvio.Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\MS" & "\immagine1.gif"
ElseIf Data = "animazione2" Then
 esegui_animazioni_MSinvio.Show
 esegui_animazioni_MSinvio.Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\MS" & "\immagine2.gif"
ElseIf Data = "animazione3" Then
 esegui_animazioni_MSinvio.Show
 esegui_animazioni_MSinvio.Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\MS" & "\immagine3.gif"
ElseIf Data = "animazione4" Then
 esegui_animazioni_MSinvio.Show
 esegui_animazioni_MSinvio.Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\MS" & "\immagine4.gif"
End If
 Listms.AddItem Data
End Sub

Private Sub WsMSinvio_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ' in caso la conessione non vada a buon fine un messaggio ci avvisa'
    errore.Show
    errore.Labelerrore.Caption = "non e' possibile effettuare la connessione al server....."
End Sub

' in questo caso il comando chiudi del form non portera' alla chiusura, verra' bloccato usando cancel = 1 '
' ed il form viene reso non visibile '
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 Cmddisconnect_Click
End Sub
