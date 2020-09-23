VERSION 5.00
Begin VB.Form avviso 
   Caption         =   "avviso"
   ClientHeight    =   1455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmdok 
      Caption         =   "ok"
      Height          =   255
      Left            =   5040
      TabIndex        =   0
      Top             =   960
      Width           =   495
   End
   Begin VB.Timer Timer_unload 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   240
      Top             =   840
   End
   Begin VB.Label Labelavviso 
      BackStyle       =   0  'Transparent
      Caption         =   "avviso :"
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
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Labelmessaggio 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   360
      Width           =   4215
   End
End
Attribute VB_Name = "avviso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdexit_Click()
cmdOk_Click
End Sub

Private Sub cmdOk_Click()
Unload avviso
End Sub

Private Sub Timer_unload_Timer()
 cmdOk_Click
End Sub

 Private Sub Form_Load()
 Timer_unload.Enabled = True
 End Sub

