VERSION 5.00
Begin VB.Form account 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      BackColor       =   &H80000013&
      Height          =   2295
      Left            =   840
      TabIndex        =   21
      Top             =   5760
      Width           =   9375
      Begin client.CandyButton Cmdcontrolla 
         Height          =   375
         Left            =   5040
         TabIndex        =   23
         Top             =   1560
         Width           =   855
         _ExtentX        =   1508
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
         Caption         =   "controlla"
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
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1560
         TabIndex        =   22
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "qual'e' il tuo piatto preferito?"
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
         Left            =   1560
         TabIndex        =   25
         Top             =   960
         Width           =   3135
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   9360
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "rispondi alla domanda per recuperare la password"
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
         Left            =   1560
         TabIndex        =   24
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000013&
      Height          =   2295
      Left            =   960
      TabIndex        =   16
      Top             =   1560
      Width           =   8895
      Begin client.CandyButton Crecuper 
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1920
         Width           =   3375
         _ExtentX        =   5953
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
         Caption         =   "recupera password"
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
      Begin client.CandyButton Cmdverifica 
         Height          =   375
         Left            =   5760
         TabIndex        =   19
         Top             =   840
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
         Caption         =   "verifica"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Style           =   2
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
         Height          =   495
         Left            =   1440
         TabIndex        =   18
         Top             =   720
         Width           =   3735
      End
      Begin client.chameleonButton chameleonButton2 
         Height          =   975
         Left            =   1200
         TabIndex        =   17
         Top             =   480
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   1720
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
         MICON           =   "account.frx":0000
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
   Begin VB.Frame Frame3 
      BackColor       =   &H000000C0&
      ForeColor       =   &H8000000D&
      Height          =   2295
      Left            =   5640
      TabIndex        =   3
      Top             =   1560
      Width           =   4215
      Begin VB.TextBox txtPassword2 
         BackColor       =   &H8000000B&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtUsername2 
         BackColor       =   &H8000000B&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "secondo account"
         Height          =   255
         Left            =   1560
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "password:"
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "username:"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000013&
      Height          =   2295
      Left            =   960
      TabIndex        =   2
      Top             =   1560
      Width           =   4215
      Begin VB.TextBox txtUsername 
         BackColor       =   &H8000000A&
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   6
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtPassword 
         BackColor       =   &H8000000B&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   12
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "primo account"
         Height          =   255
         Left            =   1680
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   1320
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   9975
      Begin VB.Timer Timer_unload 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   9120
         Top             =   3240
      End
      Begin client.CandyButton CandyButton1 
         Height          =   3855
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   6800
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
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
   End
   Begin client.chameleonButton chameleonButton1 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   8281
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
      MICON           =   "account.frx":001C
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
      Left            =   9960
      TabIndex        =   13
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
   Begin VB.Image Image10 
      Height          =   450
      Left            =   240
      Picture         =   "account.frx":0038
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10125
   End
   Begin VB.Image Image11 
      Height          =   435
      Left            =   10320
      Picture         =   "account.frx":0FF2
      Top             =   15
      Width           =   300
   End
   Begin VB.Image Image9 
      Height          =   450
      Left            =   0
      Picture         =   "account.frx":1700
      Top             =   0
      Width           =   345
   End
End
Attribute VB_Name = "account"
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

Private Sub Cmdcontrolla_Click()
 If Text2.Text = psw_account.Text1.Text Then
    Text1.Text = psw_account.Txtpsw_account.Text
    avviso.Show
    avviso.Labelmessaggio.Caption = " la risposta e' corretta "
    Frame5.Top = 5760
    Text1.Locked = True
 Else
    errore.Show
    errore.Labelerrore.Caption = " la risposta non e' corretta "
 End If
End Sub

Private Sub Cmdverifica_Click()
 If Text1.Text = psw_account.Txtpsw_account.Text Then
    Frame4.Top = 5760
    Frame4.Left = 840
 Else
    errore.Show
    errore.Labelerrore.Caption = " la password e' errata"
 End If
End Sub

Private Sub Crecuper_Click()
 Frame5.Top = 1560
 Frame5.Left = 960
End Sub

Private Sub form_load()
  Frame4.Top = 1560
  Frame4.Left = 960
  txtUsername = RegLoad(txtUsername)
  Txtpassword = RegLoad(Txtpassword)
  txtUsername2 = RegLoad(txtUsername2)
  txtPassword2 = RegLoad(txtPassword2)
  Text1.Locked = False
End Sub
Private Sub SaveControlValues()
 Call RegSave(txtUsername, txtUsername.Text)
 Call RegSave(Txtpassword, Txtpassword.Text)
 Call RegSave(txtUsername2, txtUsername.Text)
 Call RegSave(txtPassword2, Txtpassword.Text)
End Sub

Private Sub Cmdexit_Click()
 Call SaveControlValues
 Timer_unload.Enabled = True
End Sub

Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 ReleaseCapture
 SendMessage account.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub

Private Sub Timer_unload_Timer()
 Unload account
 Timer_unload.Enabled = False
End Sub

