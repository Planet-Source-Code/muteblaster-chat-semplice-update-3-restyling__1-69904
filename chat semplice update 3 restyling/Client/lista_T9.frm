VERSION 5.00
Begin VB.Form lista_T9 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox ListaT9 
      Height          =   5520
      Left            =   1080
      TabIndex        =   1
      Top             =   2040
      Width           =   2535
   End
   Begin client.chameleonButton chameleonButton1 
      Height          =   6375
      Left            =   360
      TabIndex        =   0
      Top             =   1560
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   11245
      BTYPE           =   3
      TX              =   "chameleonButton1"
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
      MICON           =   "lista_T9.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label_INFORMAZIONI 
      BackStyle       =   0  'Transparent
      Caption         =   "questa lista carica  le parole per facilitarne la scrittura......usando un qualcosa di simile al T9  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   3735
   End
End
Attribute VB_Name = "lista_T9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ListaT9_Click()
chat.Txtsend = ListaT9
End Sub
