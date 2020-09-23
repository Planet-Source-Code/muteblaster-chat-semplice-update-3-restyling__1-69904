VERSION 5.00
Begin VB.Form link 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txtrimuovilink 
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   6720
      Width           =   3615
   End
   Begin client.CandyButton Cmdrimuovi 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   6720
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "rimuovi link"
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
   Begin client.CandyButton Cmdinserisci 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "inserisci link"
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
   Begin VB.TextBox Txtlink 
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   840
      Width           =   3615
   End
   Begin VB.ListBox Listalink 
      Height          =   4545
      ItemData        =   "link.frx":0000
      Left            =   480
      List            =   "link.frx":0002
      TabIndex        =   4
      Top             =   1680
      Width           =   4335
   End
   Begin client.CandyButton cmdchiudi 
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   0
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
      Caption         =   "x"
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
   Begin client.CandyButton Cmdsalva 
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   7440
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "salva"
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
   Begin client.chameleonButton chameleonButton1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   9340
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "link.frx":0004
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Shape Shape1 
      Height          =   8055
      Left            =   0
      Top             =   0
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "memorizza i link "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "link"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdchiudi_Click()
 chat.Picture21.Top = 10800
End Sub

Private Sub Cmdinserisci_Click()
 Listalink.AddItem Txtlink
 Txtlink.Text = ""
End Sub

Private Sub Cmdrimuovi_Click()
 Dim I As Integer 'dichiariamo la variabile'
For I = (Listalink.ListCount - 1) To 0 Step -1

      If InStr(Listalink.List(I), Txtrimuovilink.Text) <> 0 Then 'mettiamo in un txt apparte l'utente selezionato'
      Listalink.RemoveItem I 'rimuoviamo l'utente'
      End If
  Next
  Txtrimuovilink.Text = ""
End Sub

Private Sub Cmdsalva_Click()
  Dim savelist As Long
    On Error Resume Next
    Open App.Path & "\link\savelist.txt" For Output As #1
    For savelist& = 0 To Listalink.ListCount - 1
        Print #1, Listalink.List(savelist&)
    Next savelist&
    Close #1
    Call MsgBox("lista salvata", vbOKOnly, List)
End Sub

Private Sub form_load()
  Call AddHScroll(Listalink)
  EnableURLDetect link.hWnd, Me.hWnd
 Dim MyString As String
    On Error Resume Next
    Open App.Path & "\link\savelist.txt" For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
       If Len(MyString$) = 0 Or MyString$ = " " Then
       
       GoTo 10
       End If
       
        Listalink.AddItem MyString$
        Listalink.Refresh
10
    Wend
20
    Close #1
End Sub

Private Sub Listalink_Click()
  chat.txtsend = Listalink.Text
End Sub

Private Sub Listalink_dblClick()
  Txtrimuovilink = Listalink.Text
End Sub
