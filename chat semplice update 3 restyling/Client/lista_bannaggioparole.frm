VERSION 5.00
Begin VB.Form lista_bannaggioparole 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "lista di parole bannate"
   ClientHeight    =   6375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3585
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List_bannaggioparole 
      Enabled         =   0   'False
      Height          =   3960
      ItemData        =   "lista_bannaggioparole.frx":0000
      Left            =   600
      List            =   "lista_bannaggioparole.frx":0007
      TabIndex        =   2
      Top             =   2040
      Width           =   2295
   End
   Begin client.chameleonButton chameleonButton1 
      Height          =   4575
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   8070
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
      MICON           =   "lista_bannaggioparole.frx":0012
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label_spiegazione 
      BackStyle       =   0  'Transparent
      Caption         =   "qui' vengono annotate tutte le parole da bannare, per il buon gusto della chat "
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
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   2775
   End
End
Attribute VB_Name = "lista_bannaggioparole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
