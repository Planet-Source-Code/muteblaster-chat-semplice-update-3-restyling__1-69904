VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form PMserverinvio 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "PM server"
   ClientHeight    =   4125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin client.CandyButton Cmdchiudi 
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   0
      Width           =   495
      _ExtentX        =   873
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
      Caption         =   "x"
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
   Begin client.CandyButton Cmdsend 
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   3240
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
      Left            =   1800
      TabIndex        =   8
      Top             =   3240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
      Left            =   480
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3615
      Begin VB.TextBox Txtnickamico 
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Labelmessaggio 
         BackStyle       =   0  'Transparent
         Caption         =   "stai scrivendo unmessaggio a"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
   End
   Begin RichTextLib.RichTextBox Txtsend 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   2990
      _Version        =   393217
      TextRTF         =   $"PMserverinvio.frx":0000
   End
   Begin MSWinsockLib.Winsock WsPMserverinvio 
      Left            =   120
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Shape Shape1 
      Height          =   4095
      Left            =   0
      Top             =   0
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "INVIA UN MESSAGGIO AL SERVER"
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
      TabIndex        =   6
      Top             =   360
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
      Left            =   240
      TabIndex        =   5
      Top             =   3720
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
      Top             =   3720
      Width           =   1335
   End
End
Attribute VB_Name = "PMserverinvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdchiudi_Click()
 Cmddisconnect_Click
End Sub

Private Sub form_load()
WsPMserverinvio.connect login.txtIP.Text, "3338"  'connetti al server'
Cmdsend.Enabled = True 'abilita il comando send'
Txtnickamico.Text = " ADMIN "
End Sub
Private Sub cmdCancel_Click()
Txtsend.Text = ""
End Sub

Private Sub Cmddisconnect_Click()
 WsPMserverinvio.Close
 Txtsend.Enabled = True
 Txtsend.Text = ""
 Unload PMserverinvio
End Sub

Private Sub WsPMserverinvio_connect()
 avviso.Labelmessaggio.Caption = "puoi spedire il messaggio al server"
End Sub

Private Sub cmdsend_Click()
WsPMserverinvio.SendData " < " & chat.Txtmionick.Text & " > " & " si e' connesso per spedirti un messaggio privato " & vbCrLf & " ---------------" & vbCrLf
If WsPMserverinvio.State = sckConnected Then
    WsPMserverinvio.SendData chat.Txtmionick.Text & " > " & Txtsend.Text
End If
Txtsend.Enabled = False
Labelinvio.Caption = chat.Label1.Caption
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 Cmddisconnect_Click
End Sub

Private Sub WsPMserverinvio_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ' se c'e' un problea nella connessione ci viene segnalato'
    MsgBox "non e' possibile eseguire la connessione al server....."
    Cmddisconnect.Enabled = False
End Sub

