VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form server_MOD_chat 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "server"
   ClientHeight    =   6630
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   10650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmdchiudi 
      Caption         =   "x"
      Height          =   255
      Left            =   10080
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin RichTextLib.RichTextBox txtchat 
      Height          =   4215
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   7435
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      TextRTF         =   $"server_MOD_chat.frx":0000
   End
   Begin VB.Frame Frameonline 
      BackColor       =   &H80000013&
      Caption         =   "utenti in linea"
      Height          =   4575
      Left            =   8040
      TabIndex        =   4
      Top             =   1200
      Width           =   2175
      Begin VB.ListBox listonline 
         Height          =   3960
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.TextBox txtNick 
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "ADMIN"
      Top             =   720
      Width           =   1635
   End
   Begin VB.CommandButton Cmdlisten 
      Caption         =   "listen"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   315
      Left            =   6840
      TabIndex        =   1
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox txtSend 
      Height          =   765
      Left            =   240
      TabIndex        =   0
      Top             =   5760
      Width           =   5955
   End
   Begin MSWinsockLib.Winsock WS 
      Index           =   0
      Left            =   9960
      Top             =   6000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Shape Shape1 
      Height          =   6615
      Left            =   0
      Top             =   0
      Width           =   10575
   End
   Begin VB.Image Image9 
      Height          =   450
      Left            =   0
      Picture         =   "server_MOD_chat.frx":0082
      Top             =   0
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image11 
      Height          =   435
      Left            =   10320
      Picture         =   "server_MOD_chat.frx":0934
      Top             =   15
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Image10 
      Height          =   450
      Left            =   240
      Picture         =   "server_MOD_chat.frx":1042
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   10125
   End
   Begin VB.Label Labelnick 
      BackStyle       =   0  'Transparent
      Caption         =   "nick dek server"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "server_MOD_chat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public ClientIndex As Integer

Private Sub Cmdchiudi_Click()
 server_MOD_chat.Visible = False
End Sub

Private Sub Cmdlisten_Click()
 Ws(0).LocalPort = 4000
 Ws(0).Listen ' in attesa di connessioni'
 listonline.AddItem "ADMIN"
End Sub

Private Sub cmdSend_Click()
Dim I As Integer

If Not txtSend.Text = "" Then
For I = 1 To ClientIndex - 1 ' se ci sono 9 client connessi il clientindex e' 10 "
' bastera' impostare il clientindex a -1 per sare il numero di client'
    txtchat.Text = txtchat.Text + " <" & txtNick.Text & "> " & txtSend.Text + vbCrLf
    Ws(I).SendData " <" & txtNick.Text & "> " & txtSend.Text ' inviamo il nostro testo al client'
Next
 txtSend.Text = "" ' dopo aver spedito il messaggio il txtsend ritorna vuoto'
End If
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
Dim I As Integer
If KeyAscii = 13 Then ' se si preme enter'
    KeyAscii = 0
    If Not txtSend.Text = "" Then
For I = 1 To ClientIndex - 1 ' se ci sono 9 client connessi il clientindex e' 10 "
' bastera' impostare il clientindex a -1 per sare il numero di client'
    txtchat.Text = txtchat.Text + " <" & txtNick.Text & "> " & txtSend.Text + vbCrLf
    Ws(I).SendData " <" & txtNick.Text & "> " & txtSend.Text ' inviamo il nostro testo al client'
Next
    End If
    txtSend.Text = "" ' dopo aver spedito il messaggio il txtsend ritorna vuoto'
End If
End Sub

Private Sub Form_Load()
ClientIndex = 1
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
    
For I = 1 To ClientIndex - 1
    Ws(I).SendData txt ' richiamo della variabile txt'
Next
    txtchat.Text = txtchat.Text + txt + vbCrLf
End Sub

Sub InviaLista()
For ii = 1 To ClientIndex - 1
    Ws(ii).SendData "@LISTC"
    For I = 0 To listonline.ListCount - 1
        Ws(ii).SendData "@LIST:" & listonline.List(I)
    Next I
Next ii

End Sub

Private Sub WS_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
 WS_Close Index
End Sub
