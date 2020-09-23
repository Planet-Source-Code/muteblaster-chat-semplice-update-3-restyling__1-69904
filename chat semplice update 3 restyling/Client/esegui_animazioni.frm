VERSION 5.00
Begin VB.Form esegui_animazioni_MSricevi 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2445
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   2445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin client.Anim Anim1 
      Height          =   1935
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   3413
   End
   Begin client.chameleonButton chameleonButton1 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   4260
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
      MICON           =   "esegui_animazioni.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer Timer_durata 
      Interval        =   3000
      Left            =   360
      Top             =   2160
   End
End
Attribute VB_Name = "esegui_animazioni_MSricevi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 Timer_durata.Enabled = True
End Sub
Private Sub Timer_durata_Timer()
Unload esegui_animazioni_MSricevi
End Sub
