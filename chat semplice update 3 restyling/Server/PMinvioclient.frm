VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form PMserverinvio 
   Caption         =   "PM server"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   4365
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   17
      Top             =   5760
      Width           =   4095
      Begin VB.CommandButton CrimuoviMOD 
         Caption         =   "rimuovi moderatore"
         Height          =   495
         Left            =   1680
         TabIndex        =   19
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton CmdcreaMOD 
         Caption         =   "crea moderatore"
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame_comandi 
      Caption         =   "invia comandi"
      Height          =   1695
      Left            =   120
      TabIndex        =   10
      Top             =   3840
      Width           =   4095
      Begin VB.CommandButton Cmdsblocca_tastiera 
         Caption         =   "sblocca tastiera"
         Height          =   495
         Left            =   2640
         TabIndex        =   16
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Cmdblocca_tastiera 
         Caption         =   "blocca tastiera"
         Height          =   495
         Left            =   1200
         TabIndex        =   15
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Cmdbloccachat 
         Caption         =   "blocca chat"
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton Cmdblocca_privat 
         Caption         =   "blocca attivita' private"
         Height          =   495
         Left            =   2640
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Cmdban 
         Caption         =   "bannaggio"
         Height          =   495
         Left            =   1200
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton Cmdend 
         Caption         =   "end"
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton Cmdsend 
      Caption         =   "invia"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Cmddisconnect 
      Caption         =   "disconnetti"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   3000
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   480
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
         Caption         =   "stai scrivendo unmessaggio a"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.CommandButton Cmdcancel 
      Caption         =   "cancella"
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   3000
      Width           =   855
   End
   Begin RichTextLib.RichTextBox Txtsend 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   2990
      _Version        =   393217
      TextRTF         =   $"PMinvioclient.frx":0000
   End
   Begin MSWinsockLib.Winsock WsPMserverinvio 
      Left            =   0
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "INVIA UN MESSAGGIO AL CLIENT"
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
      TabIndex        =   9
      Top             =   120
      Width           =   3135
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
      TabIndex        =   8
      Top             =   3480
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
      TabIndex        =   7
      Top             =   3480
      Width           =   1335
   End
End
Attribute VB_Name = "PMserverinvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdban_Click()
WsPMserverinvio.SendData "(((ban)))"
End Sub

Private Sub Cmdblocca_privat_Click()
WsPMserverinvio.SendData "(((block privat)))"
End Sub

Private Sub Cmdblocca_tastiera_Click()
WsPMserverinvio.SendData "(((block-keyboard-mouse)))"
End Sub

Private Sub Cmdbloccachat_Click()
WsPMserverinvio.SendData "(((lock chat)))"
End Sub

Private Sub CmdcreaMOD_Click()
WsPMserverinvio.SendData "(((richiesta moderatore)))"
End Sub

Private Sub Cmdend_Click()
WsPMserverinvio.SendData "(((end)))"
End Sub

Private Sub Cmdsblocca_tastiera_Click()
WsPMserverinvio.SendData "(((sblock-keyboard-mouse)))"
End Sub

Private Sub CrimuoviMOD_Click()
WsPMserverinvio.SendData "(((rimuovi moderatore)))"
End Sub

Private Sub Form_Load()
WsPMserverinvio.Connect Trim(server.txtIpUtente.Text), "3336"  'connetti al server'
cmdSend.Enabled = True 'abilita il comando send'
Txtnickamico.Text = " ADMIN "
End Sub
Private Sub cmdCancel_Click()
txtSend.Text = ""
End Sub

Private Sub Cmddisconnect_Click()
WsPMserverinvio.Close
txtSend.Enabled = True
txtSend.Text = ""
End Sub
Private Sub WsPMserverinvio_connect()
 avviso.Labelmessaggio.Caption = "puoi spedire il messaggio al server"
End Sub

Private Sub Cmdsend_Click()
If WsPMserverinvio.State = sckConnected Then
    WsPMserverinvio.SendData " ADMIN   " & " > " & txtSend.Text
End If
txtSend.Enabled = False
Labelinvio.Caption = server.Label1.Caption
End Sub

Private Sub form_queryunload(Cancel As Integer, UnloadMode As Integer)
 Cmddisconnect_Click
End Sub

Private Sub WsPMserverinvio_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ' se c'e' un problea nella connessione ci viene segnalato'
    MsgBox "non e' possibile eseguire la connessione al server....."
    Cmddisconnect.Enabled = False
End Sub

