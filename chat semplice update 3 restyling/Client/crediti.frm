VERSION 5.00
Begin VB.Form crediti 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   Caption         =   "crediti"
   ClientHeight    =   8445
   ClientLeft      =   5445
   ClientTop       =   1350
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   Begin client.CandyButton CandyButton1 
      Height          =   255
      Left            =   8880
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
      Caption         =   "x"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Height          =   7215
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   8415
      Begin VB.Frame Frame_crediti 
         BackColor       =   &H8000000D&
         Height          =   7215
         Left            =   -240
         TabIndex        =   3
         Top             =   0
         Width           =   8655
         Begin VB.PictureBox Picture4 
            Height          =   1455
            Left            =   5760
            Picture         =   "crediti.frx":0000
            ScaleHeight     =   1395
            ScaleWidth      =   1515
            TabIndex        =   10
            Top             =   1560
            Width           =   1575
         End
         Begin VB.PictureBox Picture1 
            Height          =   1695
            Left            =   1440
            Picture         =   "crediti.frx":0DEA
            ScaleHeight     =   1635
            ScaleWidth      =   1515
            TabIndex        =   6
            Top             =   960
            Width           =   1575
         End
         Begin VB.PictureBox Picture2 
            Height          =   1455
            Left            =   1440
            Picture         =   "crediti.frx":38A3
            ScaleHeight     =   1395
            ScaleWidth      =   1515
            TabIndex        =   5
            Top             =   5640
            Width           =   1575
         End
         Begin VB.PictureBox Picture3 
            Height          =   1575
            Left            =   1680
            Picture         =   "crediti.frx":7895
            ScaleHeight     =   1515
            ScaleWidth      =   1035
            TabIndex        =   4
            Top             =   3240
            Width           =   1095
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "grazie xaxak per gli smile di grandi dimensioni"
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
            Left            =   5040
            TabIndex        =   12
            Top             =   3360
            Width           =   3015
         End
         Begin VB.Image Image1 
            Height          =   675
            Left            =   5640
            Picture         =   "crediti.frx":861D
            Top             =   4080
            Width           =   1560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "grazie a marcostraf per il supporto tecnico e per l'aiuno nel risolvere importanti bug nel server"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   5040
            TabIndex        =   11
            Top             =   240
            Width           =   3375
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000000&
            X1              =   4800
            X2              =   4800
            Y1              =   120
            Y2              =   7200
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "muteblaster ideatore e sviluppatore"
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
            Left            =   240
            TabIndex        =   9
            Top             =   600
            Width           =   4335
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "ringraziamenti a roby66 per la divisione delle stringhe"
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
            Left            =   360
            TabIndex        =   8
            Top             =   2760
            Width           =   4455
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "ringraziamenti a h2201 per la lista degli utenti e gli smile"
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
            Left            =   360
            TabIndex        =   7
            Top             =   4920
            Width           =   4095
         End
      End
      Begin client.chameleonButton chameleonButton2 
         Height          =   7215
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   12726
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
         MICON           =   "crediti.frx":8F64
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
   Begin client.chameleonButton chameleonButton1 
      Height          =   7695
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   13573
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
      MICON           =   "crediti.frx":8F80
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer Timer1 
      Interval        =   40
      Left            =   0
      Top             =   6720
   End
   Begin VB.Shape Shape1 
      Height          =   8415
      Left            =   0
      Top             =   0
      Width           =   9615
   End
   Begin VB.Image Image9 
      Height          =   450
      Left            =   0
      Picture         =   "crediti.frx":8F9C
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image11 
      Height          =   435
      Left            =   9240
      Picture         =   "crediti.frx":984E
      Top             =   15
      Width           =   300
   End
   Begin VB.Image Image10 
      Height          =   450
      Left            =   240
      Picture         =   "crediti.frx":9F5C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9045
   End
End
Attribute VB_Name = "crediti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Private OldX As Integer
Private OldY As Integer

Private Sub CandyButton1_Click()
 crediti.Visible = False
End Sub

Private Sub Form_Load()
 MakeTransparent Me.hwnd, 200
End Sub

Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage crediti.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub

Private Sub Timer1_Timer()
    Frame_crediti.Top = IIf(Frame_crediti.Top <= -Frame_crediti.Height, Height, Frame_crediti.Top - 25)
End Sub
