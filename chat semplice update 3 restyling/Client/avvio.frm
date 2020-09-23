VERSION 5.00
Begin VB.Form avvio 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   ClientHeight    =   4245
   ClientLeft      =   3570
   ClientTop       =   2985
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer_caricamento_avvio 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   6960
      Top             =   2040
   End
   Begin VB.Timer Timer_verifica_avvio 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   6960
      Top             =   2520
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   6960
      Top             =   3000
   End
   Begin client.CandyButton CandyButton3 
      Height          =   495
      Left            =   6840
      TabIndex        =   5
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "X"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   6
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   2760
      Picture         =   "avvio.frx":0000
      ScaleHeight     =   1635
      ScaleWidth      =   1875
      TabIndex        =   4
      Top             =   1200
      Width           =   1935
   End
   Begin client.CandyButton CandyButton2 
      Height          =   2175
      Left            =   2280
      TabIndex        =   3
      Top             =   960
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   3836
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "CandyButton2"
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
   Begin client.CandyButton CandyButton1 
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   3360
      Width           =   1695
      _ExtentX        =   2990
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
      Caption         =   "leggi licenza"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   6
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000013&
      Caption         =   "non mostrare all'avvio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      Height          =   4215
      Left            =   0
      Top             =   0
      Width           =   7695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "l'utilizzo di questo software e' fatto nei termini della licenza gnu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   360
      Width           =   4815
   End
End
Attribute VB_Name = "avvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CandyButton2_Click()
 CandyButton1_Click
End Sub

Private Sub CandyButton3_Click()
 Unload avvio
 login.Visible = True
End Sub

Private Sub Form_Load()
  Timer_caricamento_avvio.Enabled = True
End Sub
Private Sub CandyButton1_Click()
 licenza.Show
End Sub

Private Sub Check1_Click()
 If Check1 = 1 Then
    opzioni.Check2 = 1
 End If
End Sub
' questo timer mi permettera' di caricare le dimensioni del form login quando'
' e' stato chiuso
Private Sub Timer_caricamento_avvio_Timer()
 dimensioni_form.Text_login_height = RegLoad(dimensioni_form.Text_login_height)
 dimensioni_form.Text_login_widh = RegLoad(dimensioni_form.Text_login_widh)
 dimensioni_form.Text1 = RegLoad(dimensioni_form.Text1)
 dimensioni_form.Text2 = RegLoad(dimensioni_form.Text2)
 Timer_verifica_avvio.Enabled = True
 Timer_caricamento_avvio.Enabled = False
End Sub

Private Sub Timer_verifica_avvio_Timer()
 If opzioni.Check2 = 1 Then
    avvio.Visible = False
    login.Height = dimensioni_form.Text_login_height.Text
    login.Width = dimensioni_form.Text_login_widh.Text
    login.Top = dimensioni_form.Text1.Text
    login.Left = dimensioni_form.Text2.Text
    login.Visible = True
 ElseIf opzioni.Check2 = 0 Then
    Timer1.Enabled = True
 End If
 Timer_verifica_avvio.Enabled = False
End Sub

Private Sub Timer1_Timer()
 CandyButton3_Click
 Timer1.Enabled = False
End Sub
