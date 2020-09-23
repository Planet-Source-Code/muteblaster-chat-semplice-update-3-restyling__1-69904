VERSION 5.00
Begin VB.Form riabilita_privat 
   Caption         =   "riabilita funzioni extrachat"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timerunload 
      Enabled         =   0   'False
      Interval        =   50000
      Left            =   120
      Top             =   2520
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "ok"
      Height          =   375
      Left            =   6480
      TabIndex        =   7
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton CmdriabilitaRICEVIFILE 
      Caption         =   "riabilita ricevifile"
      Height          =   495
      Left            =   4440
      TabIndex        =   5
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton CmdriabilitaPM 
      Caption         =   "riabilita PM"
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton CmdriabilitaMS 
      Caption         =   "riabilita MS"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "riabilita la possibilita' di ricevere file"
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "riabilita i singoli messaggi"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "riabilita la chat private "
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "potete riabilitare tutte quelle funzioni extrachat che in precedenza erano state disabilitate"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
   End
End
Attribute VB_Name = "riabilita_privat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 Timerunload.Enabled = True
If login.WsMSricevi.State = sckListening Then
   CmdriabilitaMS.Enabled = False
End If
If login.WsPMricevi.State = sckListening Then
   CmdriabilitaPM.Enabled = False
End If
If login.Wsricevifile.State = sckListening Then
   CmdriabilitaRICEVIFILE.Enabled = False
End If
End Sub

Private Sub CmdriabilitaPM_Click()
On Error Resume Next
login.WsPMricevi.Listen
End Sub

Private Sub CmdriabilitaMS_Click()
On Error Resume Next
login.WsMSricevi.Listen
End Sub

Private Sub CmdriabilitaRICEVIFILE_Click()
On Error Resume Next
login.Wsricevifile.Listen
End Sub

Private Sub cmdOk_Click()
Unload riabilita_privat
End Sub

Private Sub Timerunload_Timer()
If riabilita_privat.Visible = True Then
  cmdOk_Click
 End If
End Sub
