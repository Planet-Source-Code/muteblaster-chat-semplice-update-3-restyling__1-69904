VERSION 5.00
Begin VB.Form listabannati 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "lista bannati"
   ClientHeight    =   6150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin client.CandyButton Cmdripristina 
      Height          =   375
      Left            =   2280
      TabIndex        =   12
      Top             =   5520
      Width           =   1095
      _ExtentX        =   1931
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
      Caption         =   "ripristina"
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
   Begin client.CandyButton cmdcancella 
      Height          =   855
      Left            =   2280
      TabIndex        =   11
      Top             =   3000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "cancella lista"
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
   Begin client.CandyButton Cmdesci 
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   840
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
      Caption         =   "esci"
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
   Begin client.CandyButton Cmdpassa 
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "inserisci"
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
   Begin VB.TextBox Txtnick 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton Cmdload 
      Caption         =   "carica lista bannati"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   6480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Cmdsalva 
      Caption         =   "salva lista bannati"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   6240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Txtrecordbannati 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   5640
      Width           =   1815
   End
   Begin VB.ListBox bannedlist 
      BackColor       =   &H80000000&
      Height          =   2985
      Left            =   240
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "crea la tua lista di bannati"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Labelcancella 
      BackStyle       =   0  'Transparent
      Caption         =   "camcella tutta la lista bannati"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   4
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label lblripristino 
      BackStyle       =   0  'Transparent
      Caption         =   "togli l'utente dalla lista dei bannati"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "lista bannati"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   1800
      Width           =   975
   End
End
Attribute VB_Name = "listabannati"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdload_Click()
Dim MyString As String
    On Error Resume Next

    Open App.Path & "\log\liste\bannati\savelist.dat" For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
       If Len(MyString$) = 0 Or MyString$ = " " Then
       
       GoTo 10
       End If
       
        bannedlist.AddItem MyString$
        bannedlist.Refresh
10
    Wend
20
    Close #1
End Sub

Private Sub Cmdsalva_Click()
Dim savelist As Long
    On Error Resume Next
    Open App.Path & "\log\liste\bannati\savelist.dat" For Output As #1
    For savelist& = 0 To bannedlist.ListCount - 1
        Print #1, bannedlist.List(savelist&)
    Next savelist&
    Close #1
End Sub

Private Sub form_load()
 Cmdload_Click
End Sub


Private Sub bannedlist_Click()
Txtrecordbannati.Text = bannedlist.Text
End Sub

Private Sub Cmdcancella_Click()
bannedlist.Clear
End Sub

Private Sub Cmdesci_Click()
 chat.Picture12.Top = 10800
 Cmdsalva_Click
 End Sub

Private Sub Cmdpassa_Click()
bannedlist.AddItem Txtnick.Text 'aggiungiamo il nome selezionato alla lista di bannati'
End Sub

Private Sub Cmdripristina_Click()
Dim i As Integer 'dichiariamo la variabile'
For i = (bannedlist.ListCount - 1) To 0 Step -1

      If InStr(bannedlist.List(i), Txtrecordbannati.Text) <> 0 Then 'mettiamo in un txt apparte l'utente selezionato'
      bannedlist.RemoveItem i 'rimuoviamo l'utente'
      End If
  Next
End Sub
