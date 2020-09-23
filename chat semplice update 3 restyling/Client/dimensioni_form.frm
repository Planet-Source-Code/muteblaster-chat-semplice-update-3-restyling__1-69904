VERSION 5.00
Begin VB.Form dimensioni_form 
   Caption         =   "dimensioni form"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "login"
      Height          =   2775
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3855
      Begin VB.Timer Timer_rileva_posizione 
         Interval        =   10000
         Left            =   3360
         Top             =   720
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   840
         TabIndex        =   8
         Text            =   "5055"
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         TabIndex        =   6
         Text            =   "765"
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox Text_login_widh 
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Text            =   "4860"
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox Text_login_height 
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Text            =   "9630"
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label4 
         Caption         =   "left"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "top"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "width"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "height"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   615
      End
   End
End
Attribute VB_Name = "dimensioni_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' questo form memorizzera' le dimensioni dei form, quando viene fatto il resize'
' e poi si ingrandisce e si vuole ritornare alla forma prima dell'ingrandimento'
' memoriza in un txt di questo form le dimensioni che poi verranno riprese'

Private Sub Text1_Change()
 If login.WindowState = 0 Then
  If Text1.Text < 300 Then
     login.Top = 300
  End If
 End If
End Sub

Private Sub Timer_rileva_posizione_Timer()
 Text_login_height.Text = login.Height
 Text_login_widh.Text = login.Width
 Text1.Text = login.Top
 Text2.Text = login.Left
 Call SaveControlValues
End Sub
Private Sub SaveControlValues()
 Call RegSave(Text_login_height, Text_login_height.Text)
 Call RegSave(Text_login_widh, Text_login_widh.Text)
 Call RegSave(Text1, Text1.Text)
 Call RegSave(Text2, Text2.Text)
 End Sub
