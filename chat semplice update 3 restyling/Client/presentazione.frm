VERSION 5.00
Begin VB.Form presentazione 
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   8220
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check_visualizza 
      Caption         =   "visualizza all'avvio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000A&
      Enabled         =   0   'False
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2880
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000A&
      Enabled         =   0   'False
      Height          =   315
      Left            =   1080
      TabIndex        =   3
      Top             =   2400
      Width           =   2415
   End
   Begin client.CandyButton cmdaccedi 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   4320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "accedi"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "verifica stato del servizio"
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
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   6960
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "crea profilo"
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
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Image Picavatar 
      Height          =   1815
      Left            =   1440
      Picture         =   "presentazione.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label label_account 
      BackStyle       =   0  'Transparent
      Caption         =   "crea account"
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
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   5520
      Width           =   1455
   End
End
Attribute VB_Name = "presentazione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub label_account_MouseMove(label As Integer, Shift As Integer, X As Single, Y As Single)
  label_account.Font.Underline = True
End Sub

 Private Sub form_mousemove(label As Integer, Shift As Integer, X As Single, Y As Single)
  label_account.Font.Underline = False
 End Sub
