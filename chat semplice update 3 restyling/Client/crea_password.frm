VERSION 5.00
Begin VB.Form crea_password_frmcreate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "crea password registrazione"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer_chiusura 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   360
      Top             =   1440
   End
   Begin client.CandyButton Cmdaccetta 
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   1440
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      Enabled         =   0   'False
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
   Begin client.CandyButton Cmdgenera 
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   1440
      Width           =   855
      _ExtentX        =   1508
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
      Caption         =   "genera"
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
   Begin VB.TextBox Txtpassword 
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "attenzione  la password poi la devi ricordare per eseguire il login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "password:"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "crea_password_frmcreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim pass As String
Dim charactere(48) As String

Private Sub Cmdaccetta_Click()
 frmCreate.Txtpassword.Text = Txtpassword.Text
 Timer_chiusura.Enabled = True
End Sub

Private Sub Form_Load()
charactere(0) = "a"
charactere(1) = "b"
charactere(2) = "c"
charactere(3) = "d"
charactere(4) = "e"
charactere(5) = "f"
charactere(6) = "g"
charactere(7) = "h"
charactere(8) = "i"
charactere(9) = "j"
charactere(10) = "k"
charactere(11) = "l"
charactere(12) = "m"
charactere(13) = "n"
charactere(14) = "o"
charactere(15) = "p"
charactere(16) = "q"
charactere(17) = "r"
charactere(18) = "s"
charactere(19) = "t"
charactere(20) = "u"
charactere(21) = "v"
charactere(22) = "w"
charactere(23) = "x"
charactere(24) = "y"
charactere(25) = "z"
charactere(26) = "0"
charactere(27) = "1"
charactere(28) = "2"
charactere(29) = "3"
charactere(30) = "4"
charactere(31) = "5"
charactere(32) = "6"
charactere(33) = "7"
charactere(34) = "8"
charactere(35) = "9"
charactere(36) = "|"
charactere(37) = ">"
charactere(38) = "<"
charactere(39) = "!"
charactere(40) = ":"
charactere(41) = ")"
charactere(42) = "("
charactere(43) = "="
charactere(44) = "+"
charactere(45) = "/"
charactere(46) = "\"
charactere(47) = "$"
charactere(48) = "*"
End Sub

Private Sub Cmdgenera_Click()
Randomize
pass = ""
i = 0
While (i <= 8)
    pass = pass & charactere(CInt(Rnd * 48))
    i = i + 1
Wend
Txtpassword.Text = pass
Clipboard.Clear
Clipboard.SetText Txtpassword.Text
End Sub

Private Sub Timer_chiusura_Timer()
 Unload crea_password_frmcreate
 Timer_chiusura.Enabled = False
End Sub

Private Sub Txtpassword_Change()
 If Txtpassword.Text = "" Then
    Cmdaccetta.Enabled = False
 Else
    Cmdaccetta.Enabled = True
 End If
End Sub
