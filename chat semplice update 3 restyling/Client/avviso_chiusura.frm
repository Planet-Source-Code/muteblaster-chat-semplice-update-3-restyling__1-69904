VERSION 5.00
Begin VB.Form avviso_chiusura 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin client.Anim Anim1 
         Height          =   1455
         Left            =   1080
         TabIndex        =   3
         Top             =   1080
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   2566
      End
      Begin client.chameleonButton chameleonButton1 
         Height          =   1935
         Left            =   720
         TabIndex        =   2
         Top             =   840
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   3413
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
         MICON           =   "avviso_chiusura.frx":0000
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
         Height          =   2895
         Left            =   0
         Top             =   0
         Width           =   4215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "e' incorso la chiusura del programma"
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
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   3975
      End
   End
End
Attribute VB_Name = "avviso_chiusura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 Anim1.AnimatedGifPath = App.Path & "\gif" & "\immagine5" & ".gif"
End Sub
