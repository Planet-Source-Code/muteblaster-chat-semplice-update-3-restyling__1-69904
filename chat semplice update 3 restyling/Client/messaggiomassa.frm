VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form messaggiomassa 
   Caption         =   "messaggio di massa"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txtip4 
      Height          =   285
      Left            =   5400
      TabIndex        =   12
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   3015
      Begin VB.Label Label1 
         Caption         =   "scrivi un singolo messaggio a piu' utenti"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.TextBox Txtip3 
      Height          =   285
      Left            =   5400
      TabIndex        =   9
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox Txtip2 
      Height          =   285
      Left            =   5400
      TabIndex        =   8
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox Txtip1 
      Height          =   285
      Left            =   5400
      TabIndex        =   7
      Top             =   960
      Width           =   1815
   End
   Begin VB.ListBox Listautenti 
      Height          =   1860
      Left            =   3360
      Style           =   1  'Checkbox
      TabIndex        =   6
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton Cmdsend 
      Caption         =   "invia"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Cmddisconnect 
      Caption         =   "disconnetti"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Cmdcancel 
      Caption         =   "cancella"
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   2880
      Width           =   855
   End
   Begin RichTextLib.RichTextBox Txtsend 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   3201
      _Version        =   393217
      TextRTF         =   $"messaggiomassa.frx":0000
   End
   Begin MSWinsockLib.Winsock Wsmessaggiomassa 
      Left            =   3360
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
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
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   2055
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
      Left            =   2400
      TabIndex        =   4
      Top             =   3360
      Width           =   1335
   End
End
Attribute VB_Name = "messaggiomassa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' o qui' viene il bello in questo tipo di messaggio voglio spedire unos tesso'
' messaggio a piu' utenti....vediam se ci riesco '

Private Sub form_load()
Dim x 'dim x'

For x = 0 To chat.Listusers.ListCount - 1

messaggiomassa.Listautenti.AddItem chat.Listusers.List(x) 'trasferiamo i dati da una lista ad unaltra'

Next x
Wsmessaggiomassa.Connect Txtip1 & Txtip2, "3333"  'connetti al server'
Cmdsend.Enabled = True 'abilita il comando send'

End Sub
Private Sub cmdCancel_Click()
Txtsend.Text = ""
End Sub

Private Sub Cmddisconnect_Click()
Wsmessaggiomassa.Close
Txtsend.Enabled = True
Txtsend.Text = ""
End Sub

Private Sub Wsmessaggiomassa_connect()
 Wsmessaggiomassa.SendData " < " & chat.Txtmionick.Text & " > " & " si e' connesso per spedirti un messaggio privato " & vbCrLf & " ---------------" & vbCrLf
End Sub

Private Sub Cmdsend_Click()
If Wsmessaggiomassa.State = sckConnected Then
    Wsmessaggiomassa.SendData "Client: " & Txtsend.Text
End If
Txtsend.Enabled = False
Labelinvio.Caption = chat.Label1.Caption
End Sub

Private Sub form_queryunload(Cancel As Integer, UnloadMode As Integer)
 Cmddisconnect_Click
End Sub

Private Sub Wsmessaggiomassa_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ' se c'e' un problea nella connessione ci viene segnalato'
    MsgBox "non e' possibile eseguire la connessione al server....."
    Cmddisconnect.Enabled = False
End Sub
