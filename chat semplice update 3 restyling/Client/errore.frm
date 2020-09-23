VERSION 5.00
Begin VB.Form errore 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "messaggio di errore"
   ClientHeight    =   2685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin client.CandyButton Cmdchiudi 
      Height          =   255
      Left            =   5760
      TabIndex        =   3
      Top             =   120
      Width           =   495
      _ExtentX        =   873
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
      Caption         =   "X"
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
   Begin VB.Timer Timer_unload 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   480
      Top             =   1920
   End
   Begin client.Anim Anim1 
      Height          =   1335
      Left            =   2400
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2355
   End
   Begin VB.Shape Shape1 
      Height          =   2655
      Left            =   0
      Top             =   0
      Width           =   6495
   End
   Begin VB.Label Labelerrore 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   5055
   End
   Begin VB.Label Labelchiamata 
      BackStyle       =   0  'Transparent
      Caption         =   "errore : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "errore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' all'avvio il timer per la chiusura si attiva'
Private Sub form_load()
  Anim1.AnimatedGifPath = App.Path & "\gif" & "\immagine4" & ".gif"
 Timer_unload.Enabled = True
End Sub

Private Sub Cmdchiudi_Click()
 Unload errore
End Sub

Private Sub Timer_unload_Timer()
 Cmdchiudi_Click
 Timer_unload.Enabled = False
End Sub
 
  
