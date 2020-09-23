VERSION 5.00
Begin VB.Form risposte 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "risposte predefinite"
   ClientHeight    =   9435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   ScaleHeight     =   9435
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timerrisposta1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   5640
      Top             =   2280
   End
   Begin VB.Frame Framerisposte 
      BackColor       =   &H00FF8080&
      Caption         =   "risposte"
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   6495
      Begin VB.Frame Frame4 
         BackColor       =   &H00FF8080&
         Caption         =   "Frase4"
         Height          =   1575
         Left            =   240
         TabIndex        =   16
         Top             =   5880
         Width           =   6015
         Begin client.CandyButton cmdrisposta4 
            Height          =   375
            Left            =   4920
            TabIndex        =   18
            Top             =   480
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
            Caption         =   "invia"
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
         Begin VB.TextBox Txtrisposta4 
            Height          =   615
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   4575
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FF8080&
         Caption         =   "Frase3"
         Height          =   1575
         Left            =   240
         TabIndex        =   12
         Top             =   4080
         Width           =   6015
         Begin client.CandyButton cmdrisposta3 
            Height          =   375
            Left            =   4920
            TabIndex        =   15
            Top             =   480
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
            Caption         =   "invia"
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
         Begin VB.TextBox Txtrisposta3 
            Height          =   615
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   4575
         End
      End
      Begin client.CandyButton Cmdok 
         Height          =   375
         Left            =   5640
         TabIndex        =   10
         Top             =   7800
         Width           =   615
         _ExtentX        =   1085
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Caption         =   "frase2"
         Height          =   1575
         Left            =   240
         TabIndex        =   4
         Top             =   2280
         Width           =   6015
         Begin VB.CheckBox Checkautmess2 
            BackColor       =   &H00FF8080&
            Caption         =   "invia in automatico ogni minuto"
            Height          =   255
            Left            =   1440
            TabIndex        =   9
            Top             =   1080
            Width           =   2535
         End
         Begin VB.Timer Timerrisposta2 
            Interval        =   60000
            Left            =   5280
            Top             =   960
         End
         Begin VB.TextBox Txtrisposta2 
            Height          =   615
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   4575
         End
         Begin client.chameleonButton cmdrisposta2 
            Height          =   375
            Left            =   4920
            TabIndex        =   5
            Top             =   480
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "invia"
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
            MICON           =   "risposte.frx":0000
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
      Begin VB.Frame Frame1 
         BackColor       =   &H00FF8080&
         Caption         =   "frase1"
         Height          =   1695
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   6015
         Begin client.CandyButton Cmdrisposta1 
            Height          =   375
            Left            =   4920
            TabIndex        =   11
            Top             =   480
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
            Caption         =   "invia"
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
         Begin VB.TextBox Txtrisposta1 
            Height          =   615
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   4695
         End
         Begin VB.CheckBox Checkautmess1 
            BackColor       =   &H00FF8080&
            Caption         =   "invia in autimatico ogni  minuto"
            Height          =   255
            Left            =   1440
            TabIndex        =   3
            Top             =   1200
            Width           =   2535
         End
      End
      Begin client.CandyButton Cmdsalva 
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   7800
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "salva"
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
   Begin client.CandyButton Cmdexit 
      Height          =   255
      Left            =   6000
      TabIndex        =   8
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
   Begin VB.Shape Shape1 
      Height          =   9375
      Left            =   0
      Top             =   0
      Width           =   6735
   End
   Begin VB.Image Image10 
      Height          =   450
      Left            =   240
      Picture         =   "risposte.frx":001C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6165
   End
   Begin VB.Image Image11 
      Height          =   435
      Left            =   6360
      Picture         =   "risposte.frx":0FD6
      Top             =   15
      Width           =   300
   End
   Begin VB.Image Image9 
      Height          =   450
      Left            =   0
      Picture         =   "risposte.frx":16E4
      Top             =   0
      Width           =   345
   End
   Begin VB.Label Labelrisposta 
      BackStyle       =   0  'Transparent
      Caption         =   "scrivi le risposte predefinite  che puoi inviare"
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
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   4815
   End
End
Attribute VB_Name = "risposte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Sub cmdrisposta3_Click()
 login.WS.SendData chat.Txtmionick.Text & "> " & "<" & login.Txtfrase.Text & ">" & vbCrLf & Txtrisposta3.Text
End Sub

Private Sub cmdrisposta4_Click()
 login.WS.SendData chat.Txtmionick.Text & "> " & "<" & login.Txtfrase.Text & ">" & vbCrLf & Txtrisposta4.Text
End Sub

Private Sub form_load()
 Txtrisposta1 = RegLoad(Txtrisposta1)
 Txtrisposta2 = RegLoad(Txtrisposta2)
 Txtrisposta3 = RegLoad(Txtrisposta3)
 Txtrisposta4 = RegLoad(Txtrisposta4)
End Sub

Private Sub SaveControlValues()
Call RegSave(Txtrisposta1, Txtrisposta1.Text)
Call RegSave(Txtrisposta2, Txtrisposta2.Text)
Call RegSave(Txtrisposta3, Txtrisposta3.Text)
Call RegSave(Txtrisposta4, Txtrisposta4.Text)
End Sub

Private Sub Checkautmess1_Click()
If Checkautmess2 = 1 Then
   Checkautmess1 = 0
ElseIf Checkautmess1 = 1 Then
Timerrisposta1.Enabled = True
 End If
End Sub

Private Sub Checkautmess2_Click()
 If Checkautmess1 = 1 Then
  Checkautmess2 = 0
 ElseIf Checkautmess2 = 1 Then
  Timerrisposta2.Enabled = True
 End If
End Sub

Private Sub Cmdexit_Click()
Cmdok_Click
End Sub

Private Sub Cmdok_Click()
 chat.Picture13.Top = 10800
End Sub

Private Sub Cmdrisposta1_Click()
login.WS.SendData chat.Txtmionick.Text & "> " & "<" & login.Txtfrase.Text & ">" & vbCrLf & Txtrisposta1.Text
End Sub

Private Sub cmdrisposta2_Click()
login.WS.SendData chat.Txtmionick.Text & "> " & "<" & login.Txtfrase.Text & ">" & vbCrLf & Txtrisposta2.Text
End Sub

Private Sub Cmdsalva_Click()
  Call SaveControlValues
End Sub

Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ReleaseCapture
SendMessage risposte.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub

Private Sub Timerrisposta1_Timer()
If Checkautmess1 = 1 Then
 Cmdrisposta1_Click
End If
End Sub



