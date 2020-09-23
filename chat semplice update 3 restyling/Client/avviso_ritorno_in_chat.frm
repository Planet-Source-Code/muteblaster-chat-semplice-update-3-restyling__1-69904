VERSION 5.00
Begin VB.Form avviso_ritorno_in_chat 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   4695
      Begin client.Anim Anim1 
         Height          =   1575
         Left            =   1200
         TabIndex        =   3
         Top             =   960
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   2778
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   405
         Left            =   960
         TabIndex        =   2
         Text            =   "BUON RITORNO IN CHAT"
         Top             =   360
         Width           =   2775
      End
      Begin client.chameleonButton chameleonButton1 
         Height          =   3135
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   5530
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
         MICON           =   "avviso_ritorno_in_chat.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
End
Attribute VB_Name = "avviso_ritorno_in_chat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 Anim1.AnimatedGifPath = App.Path & "\gif" & "\immagine2" & ".gif"
End Sub
