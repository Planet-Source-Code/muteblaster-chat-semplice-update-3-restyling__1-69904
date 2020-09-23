VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form server 
   BorderStyle     =   0  'None
   Caption         =   "server"
   ClientHeight    =   7815
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   11010
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdMOD_chat 
      Caption         =   "mostra chat moderatori"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   840
      Width           =   2055
   End
   Begin VB.Timer Timerora 
      Interval        =   1
      Left            =   9120
      Top             =   1080
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Height          =   855
      Left            =   8400
      TabIndex        =   20
      Top             =   120
      Width           =   2175
      Begin VB.Label Label2 
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
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   840
         TabIndex        =   24
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Labeldata 
         BackStyle       =   0  'Transparent
         Caption         =   "data :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   960
         TabIndex        =   22
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Labelora 
         BackStyle       =   0  'Transparent
         Caption         =   "ora :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   480
         TabIndex        =   21
         Top             =   480
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   120
         Picture         =   "server.frx":0000
         Top             =   360
         Width           =   240
      End
   End
   Begin VB.Frame Frameinformazioni 
      BackColor       =   &H80000013&
      Height          =   3495
      Left            =   8400
      TabIndex        =   14
      Top             =   1200
      Width           =   2175
      Begin VB.PictureBox Picavatar 
         BackColor       =   &H80000013&
         Height          =   1575
         Left            =   240
         ScaleHeight     =   1515
         ScaleWidth      =   1515
         TabIndex        =   17
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtIpUtente 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   360
         Width           =   1815
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000013&
         BackStyle       =   1  'Opaque
         Height          =   1815
         Left            =   120
         Shape           =   5  'Rounded Square
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Labeliputente 
         BackStyle       =   0  'Transparent
         Caption         =   "txtiputente"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Labelnickutente 
         BackStyle       =   0  'Transparent
         Caption         =   "nick"
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
         Left            =   720
         TabIndex        =   18
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.CommandButton Cmdinformazioniutente 
      Caption         =   "informazioni utente"
      Height          =   255
      Left            =   8880
      TabIndex        =   13
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   9360
      TabIndex        =   12
      Top             =   7080
      Width           =   1095
   End
   Begin VB.Frame Framerimuovi 
      Caption         =   "rimuovi utente"
      Height          =   855
      Left            =   3120
      TabIndex        =   8
      Top             =   120
      Width           =   5175
      Begin VB.TextBox Txtindiceutente 
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Txtrimuoviutente 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton Cmdrimuovi 
         Caption         =   "rimuovi utente"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
   End
   Begin RichTextLib.RichTextBox txtchat 
      Height          =   4935
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   8705
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      TextRTF         =   $"server.frx":1272
   End
   Begin VB.Frame Frameonline 
      Caption         =   "utenti in linea"
      Height          =   5055
      Left            =   6480
      TabIndex        =   4
      Top             =   1200
      Width           =   1815
      Begin VB.ListBox listonline 
         Height          =   4560
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.TextBox txtNick 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "ADMIN"
      Top             =   240
      Width           =   1635
   End
   Begin VB.CommandButton Cmdlisten 
      Caption         =   "listen"
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   315
      Left            =   5760
      TabIndex        =   1
      Top             =   7080
      Width           =   615
   End
   Begin VB.TextBox txtSend 
      Height          =   765
      Left            =   120
      TabIndex        =   0
      Top             =   6720
      Width           =   5475
   End
   Begin MSWinsockLib.Winsock WS 
      Index           =   0
      Left            =   6600
      Top             =   6960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock WsPMricevi 
      Left            =   7080
      Top             =   6960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   3338
   End
   Begin VB.Label Labelnick 
      BackStyle       =   0  'Transparent
      Caption         =   "nick dek server"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private intMax As Long
Dim scknum As Integer
Public ClientIndex As Integer

Private Sub Cmdinformazioniutente_Click()
informazioniutente.Show
End Sub

Private Sub Cmdlisten_Click()
On Error Resume Next
 Ws(0).LocalPort = 1000
 Ws(0).Listen ' in attesa di connessioni'
 server_MOD_chat.Ws(0).LocalPort = 4000
 server_MOD_chat.Ws(0).Listen ' in attesa di connessioni'
 listonline.AddItem "ADMIN" & "    " & Ws(0).LocalIP
 server_MOD_chat.listonline.AddItem "ADMIN"
End Sub

Private Sub CmdMOD_chat_Click()
 server_MOD_chat.Show
End Sub

Private Sub cmdSend_Click()
Dim I As Integer
If Not txtSend.Text = "" Then
For I = 1 To ClientIndex - 1 ' se ci sono 9 client connessi il clientindex e' 10 "
' bastera' impostare il clientindex a -1 per sare il numero di client'
    txtchat.Text = txtchat.Text & "<" & txtNick.Text & "> " & vbCrLf & txtSend.Text & vbCrLf
    Ws(I).SendData Chr(127) & "nick:" & txtNick.Text & Chr(127) & "frase:" & Chr(127) & vbCrLf & txtSend.Text ' inviamo il nostro testo al client'
    
Next
 txtSend.Text = "" ' dopo aver spedito il messaggio il txtsend ritorna vuoto'
End If
End Sub

Private Sub Cmdrimuovi_Click()
Dim I As Integer 'dichiariamo la variabile'
For I = (listonline.ListCount - 1) To 0 Step -1

      If InStr(listonline.List(I), Txtrimuoviutente.Text) <> 0 Then 'mettiamo in un txt apparte l'utente selezionato'
      listonline.RemoveItem I 'rimuoviamo l'utente'
      End If
  Next
  InviaLista
  Txtrimuoviutente.Text = ""
  Txtindiceutente.Text = ""
End Sub


Private Sub Command1_Click()
PMserverinvio.Show
End Sub

Private Sub listonline_Click()
Txtrimuoviutente.Text = listonline.Text
Txtindiceutente.Text = listonline.ListIndex
informazioniutente.txtRecord = listonline.Text
 informazioniutente.Txtutente.Text = listonline.ListIndex
 informazioniutente.txtPosizioneSpazi.Text = InStr(informazioniutente.txtRecord.Text, informazioniutente.txtInput.Text) 'viene calcolata la posizione della virgola del record
 informazioniutente.txtip.Text = Mid$(informazioniutente.txtRecord.Text, informazioniutente.txtPosizioneSpazi + 2, Left)
 informazioniutente.Text2.Text = Left$(informazioniutente.txtRecord.Text, informazioniutente.txtPosizioneSpazi.Text - 1)
 
   informazioniutente.txtPosizioneSpaziIpUtente.Text = InStr(informazioniutente.txtip.Text, informazioniutente.txtInput.Text) 'viene calcolata la posizione della virgola del record
 informazioniutente.txtIpUtente.Text = Left(informazioniutente.txtip.Text, informazioniutente.txtPosizioneSpaziIpUtente.Text) 'AAAAAAAAAAHHHHHHHHHHHH!!!!!!!!!!!! (Urlo da stress! puoi cancellarlo)
 
 informazioniutente.txtPosizioneSpaziAvatar.Text = InStr(informazioniutente.txtip.Text, informazioniutente.txtInput.Text) 'viene calcolata la posizione della virgola del record
 informazioniutente.Txtavataramico.Text = Mid$(informazioniutente.txtip.Text, informazioniutente.txtPosizioneSpaziAvatar.Text + 3, Left)  'mi mostra il numero dell'Avatar
  ' GRAZIE ROBY PER IL GRANDE AIUTO '
 Picavatar.Picture = LoadPicture(App.Path & "\avatar" & "\immagine" & informazioniutente.Txtavataramico & ".gif") ' richiamiamo l'immagine in base all'indice'
                                                                                                ' che nel form avatar si e' inserito '
 informazioniutente.Picavatar.Picture = LoadPicture(App.Path & "\avatar" & "\immagine" & informazioniutente.Txtavataramico & ".gif")
End Sub

Private Sub Timerora_Timer()
Label1.Caption = Time
Label2.Caption = Date
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
Dim I As Integer
If KeyAscii = 13 Then ' se si preme enter'
    KeyAscii = 0
    If Not txtSend.Text = "" Then
For I = 1 To ClientIndex - 1 ' se ci sono 9 client connessi il clientindex e' 10 "
' bastera' impostare il clientindex a -1 per sare il numero di client'
    txtchat.Text = txtchat.Text & "<" & txtNick.Text & "> " & vbCrLf & txtSend.Text & vbCrLf
    Ws(I).SendData Chr(127) & "nick:" & txtNick.Text & Chr(127) & "frase:" & Chr(127) & vbCrLf & txtSend.Text ' inviamo il nostro testo al client'
Next
    End If
    txtSend.Text = "" ' dopo aver spedito il messaggio il txtsend ritorna vuoto'
End If
End Sub

Private Sub Form_Load()
ClientIndex = 1
Cmdlisten_Click
WsPMricevi.Listen
End Sub

Private Sub WS_Close(Index As Integer)
ClientIndex = ClientIndex - 1
Ws(Index).Close
Unload Ws(Index)
End Sub

Private Sub Ws_ConnectionRequest(Index As Integer, ByVal requestID As Long)
If ClientIndex = 1 Then
    Load Ws(ClientIndex)
    Ws(ClientIndex).Close
    Ws(ClientIndex).Accept requestID ' accetto la connessione'
    ClientIndex = ClientIndex + 1
ElseIf ClientIndex > 1 Then
    Load Ws(ClientIndex) ' attivo un nuovo winsock'
    Ws(ClientIndex).Close ' chiudi'
    Ws(ClientIndex).Accept requestID
    ClientIndex = ClientIndex + 1
End If
End Sub

Private Sub Ws_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim txt As String, I As Integer, Sck As Object ' txt = indica cio' che si riceve'
Dim Nick, Frase, Messaggio As String

    Ws(Index).GetData txt, vbString
    
    If Mid(txt, 1, 9) = "@CONNECT:" Then
        listonline.AddItem Mid(txt, 10)
        InviaLista
        Exit Sub
    End If
     
    If Mid(txt, 1, 12) = "@DISCONNECT:" Then
        For I = 0 To listonline.ListCount - 1
            If listonline.List(I) = Mid(txt, 13) Then
                listonline.RemoveItem I
                InviaLista
            Exit Sub
            End If
        Next I
    End If
        
    If InStr(1, txt, Chr(127) & "cambionick:") > 0 Then
    server.txtchat.Text = server.txtchat.Text & Mid(txt, InStr(1, txt, Chr(127) & "nick:") + 6, InStr(1, txt, Chr(127) & "cambionick:") - (InStr(1, txt, Chr(127) & "nick:") + 6)) & Mid(txt, InStr(1, txt, Chr(127) & "cambionick:") + 12) & vbCrLf
    For I = 1 To ClientIndex - 1
    Ws(I).SendData txt ' richiamo della variabile txt'
    Next
    Exit Sub
    End If
    
For I = 1 To ClientIndex - 1
    Ws(I).SendData txt ' richiamo della variabile txt'
Next

Nick = Mid(txt, InStr(1, txt, Chr(127) & "nick:") + 6, InStr(1, txt, Chr(127) & "frase:") - (InStr(1, txt, Chr(127) & "nick:") + 6))
Frase = Mid(txt, InStr(1, txt, Chr(127) & "frase:") + 7, InStr(1, txt, Chr(127) & vbCrLf) - (InStr(1, txt, Chr(127) & "frase:") + 7))
Messaggio = Mid(txt, InStr(1, txt, Chr(127) & vbCrLf) + 3)
If Trim(Frase) = "" Then
    txtchat.Text = txtchat.Text & "<" & Nick & ">" & vbCrLf & Messaggio & vbCrLf
Else
    txtchat.Text = txtchat.Text & "<" & Nick & "> " & " <" & Frase & ">" & vbCrLf & Messaggio & vbCrLf
End If
End Sub

Private Sub WS_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
 WS_Close Index
End Sub

Sub InviaLista()
For ii = 1 To ClientIndex - 1
    Ws(ii).SendData "@LISTC"
    For I = 0 To listonline.ListCount - 1
        Ws(ii).SendData "@LIST:" & listonline.List(I)
    Next I
Next ii

End Sub

Private Sub WsPMricevi_ConnectionRequest(ByVal requestID As Long)
    If WsPMricevi.State <> sckClosed Then WsPMricevi.Close 'if state is closed, then close the socket
    WsPMricevi.Accept requestID ' accept connection from client
    PMricevi.Show
End Sub

Private Sub WsPMricevi_close()
If WsPMricevi.State <> sckClosed Then WsPMricevi.Close
 WsPMricevi.Listen
End Sub

Private Sub WsPMricevi_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
WsPMricevi.GetData Data ' store data in Data
PMricevi.txtchat.Text = Data 'add data to listbox
End Sub


