VERSION 5.00
Begin VB.Form eventi 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "eventi"
   ClientHeight    =   6015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin client.CandyButton Cmdconferma 
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   5520
      Width           =   1575
      _ExtentX        =   2778
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
      Caption         =   "<<<<<conferma"
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
   Begin client.CandyButton Cmdcancella 
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   4800
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "CANCELLA"
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
   Begin client.CandyButton Cmdload 
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   4800
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "CARICAMENTO EVENTI"
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
   Begin client.CandyButton Cmdsalva 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   4800
      Width           =   1575
      _ExtentX        =   2778
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
      Caption         =   "SALVA"
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
   Begin client.CandyButton Cmdchiudi 
      Height          =   495
      Left            =   5760
      TabIndex        =   3
      Top             =   0
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "X"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   3
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.TextBox Textevento 
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   4335
   End
   Begin VB.ListBox Listaeventi 
      BackColor       =   &H8000000A&
      Height          =   3960
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   6375
   End
   Begin VB.Label Labeleventi 
      BackStyle       =   0  'Transparent
      Caption         =   "memorizza gli eventi o le considerazioni "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "eventi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdchiudi_Click()
 chat.Picture7.Top = 10800
 Cmdsalva_Click
End Sub

Private Sub Cmdconferma_Click()
 Listaeventi.AddItem Textevento
End Sub

Private Sub form_load()
 Cmdload_Click
End Sub
Private Sub Cmdcancella_Click()
Listaeventi.Clear
End Sub

Private Sub Cmdload_Click()
Dim MyString As String
    On Error Resume Next

    Open App.Path & "\log\eventi\savelist.dat" For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
       If Len(MyString$) = 0 Or MyString$ = " " Then
       
       GoTo 10
       End If
       
        Listaeventi.AddItem MyString$
        Listaeventi.Refresh
10
    Wend
20
    Close #1
End Sub

Private Sub Cmdsalva_Click()
 Dim savelist As Long
    On Error Resume Next
    Open App.Path & "\log\eventi\savelist.dat" For Output As #1
    For savelist& = 0 To Listaeventi.ListCount - 1
        Print #1, Listaeventi.List(savelist&)
    Next savelist&
    Close #1
    Call MsgBox("lista salvata", vbOKOnly, List)
End Sub


