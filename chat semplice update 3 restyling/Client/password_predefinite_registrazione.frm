VERSION 5.00
Begin VB.Form password_predefinite_registrazione 
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2760
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin client.CandyButton Cmd_ok 
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   840
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ok"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   1620
      ItemData        =   "password_predefinite_registrazione.frx":0000
      Left            =   120
      List            =   "password_predefinite_registrazione.frx":0019
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "password_predefinite_registrazione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_ok_Click()
 frmCreate.Txtpassword.Text = Text1.Text
 Unload password_predefinite_registrazione
End Sub

Private Sub List1_Click()
 Text1.Text = List1.Text
End Sub
