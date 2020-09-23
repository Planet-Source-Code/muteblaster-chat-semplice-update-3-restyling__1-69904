VERSION 5.00
Begin VB.Form stato 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1830
   LinkTopic       =   "Form1"
   ScaleHeight     =   1275
   ScaleWidth      =   1830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer_selezione 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1320
      Top             =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "(in linea)"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Image Image18 
      Height          =   240
      Left            =   240
      Picture         =   "stato.frx":0000
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "(occupato)"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "(non al pc)"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   495
      Width           =   855
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      Height          =   255
      Left            =   120
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image Image15 
      Height          =   240
      Left            =   240
      Picture         =   "stato.frx":038A
      Top             =   840
      Width           =   240
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      Height          =   285
      Left            =   120
      Top             =   480
      Width           =   1455
   End
   Begin VB.Image Image17 
      Height          =   240
      Left            =   240
      Picture         =   "stato.frx":0714
      Top             =   480
      Width           =   240
   End
   Begin VB.Image Image16 
      Height          =   240
      Left            =   120
      Picture         =   "stato.frx":0A9E
      Top             =   1320
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "stato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Label1_Click()
 login.Label8.Caption = Label1.Caption
  Timer_selezione.Enabled = True
End Sub

Private Sub Label2_Click()
 login.Label8.Caption = Label2.Caption
 Timer_selezione.Enabled = True
End Sub

Private Sub Label3_Click()
 login.Label8.Caption = Label3.Caption
 Timer_selezione.Enabled = True
End Sub

Private Sub Timer_selezione_Timer()
 login.Picture12.Height = 9495
 login.Picture12.Width = 4815
 login.Picture12.Left = 5760
 login.Picture12.Top = 0
 Unload stato
 login.Picture12.Visible = False
 Timer_selezione.Enabled = False
End Sub
