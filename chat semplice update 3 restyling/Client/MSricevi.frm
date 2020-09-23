VERSION 5.00
Begin VB.Form MSricevi 
   BorderStyle     =   0  'None
   Caption         =   "ricevi chat privata"
   ClientHeight    =   5355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   9930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin client.CandyButton Cmdsend 
      Height          =   375
      Left            =   7080
      TabIndex        =   7
      Top             =   4680
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
   Begin client.CandyButton Cmdchiudi 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   240
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
      Caption         =   "chiudi"
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
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3960
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
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   3960
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
      Caption         =   "animazioni"
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
   Begin VB.Timer Timer_esegui_animazioni 
      Interval        =   1
      Left            =   7200
      Top             =   120
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   1680
      TabIndex        =   2
      Top             =   0
      Width           =   4695
      Begin VB.TextBox Txtmionick 
         BackColor       =   &H80000000&
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.TextBox Txtsend 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   6855
   End
   Begin VB.ListBox Listms 
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   7935
   End
End
Attribute VB_Name = "MSricevi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Newalert As New cpopup ' dichiariamo la variabile che mi permettera
                           ' di aprire nuovi popup ogni volta '
                           ' richiamando il modulo di classe cpopup'

Private Sub Cmdanimazioni_Click()
animazioni_MSricevi.Show
End Sub

Private Sub Cmdnudge_Click()
nudge.TimernudgeMSricevi.Enabled = True
End Sub

Private Sub form_load()
Txtmionick.Text = chat.Txtmionick.Text
End Sub

Private Sub Cmdchiudi_Click()
Listms.Clear
Unload MSricevi_styleE
Unload Me
End Sub

Private Sub cmdSend_Click()
If login.WsMSricevi.State = sckConnected Then 'verifica se sei connesso'
    If Txtsend.Text = "" Then ' verifica se il txtsend e' connesso'
        Else
            login.WsMSricevi.SendData (Txtmionick.Text & " > " & Txtsend.Text) 'invia il messaggio'
            Listms.AddItem Txtmionick.Text & " > " & Txtsend.Text 'metti a video il messaggio'
            Txtsend.Text = "" ' cancella il teso del txtsend'
    End If
End If
End Sub

' rendiamo inattivo il comando chiudi e' necessario farlo'
' altrimenti i winsock si chiudono'
Private Sub form_queryunload(Cancel As Integer, UnloadMode As Integer)
 Cmdchiudi_Click
End Sub


Private Sub Timer_esegui_animazioni_Timer()
esegui_animazioni_MSricevi.Top = MSricevi.Top + 3000
esegui_animazioni_MSricevi.Left = MSricevi.Left + 2000
End Sub
