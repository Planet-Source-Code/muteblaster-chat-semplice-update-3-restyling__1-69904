VERSION 5.00
Begin VB.Form psw_account 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   ClientHeight    =   2910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer_salvataggio 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6120
      Top             =   600
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   1080
      TabIndex        =   9
      Top             =   1800
      Width           =   3735
   End
   Begin VB.Timer Timer_unload 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6720
      Top             =   600
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Height          =   1815
      Left            =   0
      TabIndex        =   5
      Top             =   3120
      Width           =   7455
      Begin client.CandyButton Cmdok 
         Height          =   495
         Left            =   5640
         TabIndex        =   8
         Top             =   840
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   873
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
         Style           =   3
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
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   960
         Width           =   3495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "qual'e il tuo piatto preferito?"
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
         Left            =   1920
         TabIndex        =   6
         Top             =   240
         Width           =   3015
      End
   End
   Begin client.CandyButton Cmdsalva 
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   1800
      Width           =   975
      _ExtentX        =   1720
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
      Caption         =   "salva"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   5
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.TextBox Txtpsw_account 
      BackColor       =   &H8000000A&
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   3735
   End
   Begin client.chameleonButton chameleonButton1 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   3201
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
      MICON           =   "psw_account.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin client.CandyButton Cmdexit 
      Height          =   255
      Left            =   6840
      TabIndex        =   4
      Top             =   120
      Width           =   375
      _ExtentX        =   661
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
      Caption         =   "X"
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
   Begin VB.Image Image9 
      Height          =   450
      Left            =   0
      Picture         =   "psw_account.frx":001C
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image11 
      Height          =   435
      Left            =   7200
      Picture         =   "psw_account.frx":08CE
      Top             =   15
      Width           =   300
   End
   Begin VB.Image Image10 
      Height          =   450
      Left            =   240
      Picture         =   "psw_account.frx":0FDC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7005
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "crea password per accesso all'account"
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
      Left            =   1320
      TabIndex        =   3
      Top             =   600
      Width           =   4695
   End
End
Attribute VB_Name = "psw_account"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Private OldX As Integer
Private OldY As Integer

Private Sub cmdexit_Click()
 Timer_unload.Enabled = True
End Sub

Private Sub form_load()
 Txtpsw_account = RegLoad(Txtpsw_account)
 Text1 = RegLoad(Text1)
 If Text2.Text = "" Then
    CMDSALVA.Enabled = False
 Else
    CMDSALVA.Enabled = True
 End If
 If Text1.Text = "" Then
    Cmdok.Enabled = False
 Else
    Cmdok.Enabled = True
 End If
End Sub

Private Sub SaveControlValues()
 Call RegSave(Txtpsw_account, Txtpsw_account.Text)
End Sub

Private Sub Cmdok_Click()
 Call RegSave(Text1, Text1.Text)
 Frame1.Top = 3120
 Timer_unload.Enabled = True
End Sub

Private Sub Cmdsalva_Click()
 Txtpsw_account.Text = Text2.Text
 Timer_salvataggio.Enabled = True
End Sub

Private Sub Text1_Change()
 If Text1.Text = "" Then
     Cmdok.Enabled = False
 Else
    Cmdok.Enabled = True
 End If
End Sub

Private Sub Text2_Change()
 If Text2.Text = "" Then
    CMDSALVA.Enabled = False
 Else
    CMDSALVA.Enabled = True
 End If
End Sub

Private Sub Timer_salvataggio_Timer()
 Call SaveControlValues
 Frame1.Top = 1080
 Text2.Text = ""
 Timer_salvataggio.Enabled = False
End Sub

Private Sub Timer_unload_Timer()
 Unload psw_account
 Timer_unload.Enabled = False
End Sub

Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ReleaseCapture
SendMessage psw_account.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub
