VERSION 5.00
Begin VB.Form esegui_animazioni_MSinvio 
   BorderStyle     =   0  'None
   Caption         =   "esegui animazioni nviate"
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2310
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   2310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin client.Anim Anim1 
      Height          =   1575
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2778
   End
   Begin client.chameleonButton chameleonButton1 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   3625
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
      MICON           =   "esegui_animazioni_inviate.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer Timer_anim1 
      Interval        =   3000
      Left            =   1200
      Top             =   1560
   End
End
Attribute VB_Name = "esegui_animazioni_MSinvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer_anim1_Timer()
Unload esegui_animazioni_MSinvio
End Sub
