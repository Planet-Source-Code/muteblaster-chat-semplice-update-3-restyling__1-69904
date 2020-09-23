VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form stato_connessione 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   ScaleHeight     =   9360
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      Height          =   9375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin client.CandyButton cmdStatus 
         Height          =   495
         Left            =   1200
         TabIndex        =   3
         Top             =   4080
         Width           =   2535
         _ExtentX        =   4471
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
         Caption         =   "verifica stato connessione"
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
      Begin VB.TextBox txtStatus 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   2760
         Width           =   3135
      End
      Begin client.chameleonButton chameleonButton1 
         Height          =   9375
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   16536
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
         MICON           =   "stato_connessione.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSWinsockLib.Winsock wskNetworkStatus 
         Index           =   0
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         Protocol        =   1
      End
   End
End
Attribute VB_Name = "stato_connessione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 cmdStatus_Click
End Sub

Private Sub cmdStatus_Click()
   LoadNewWinsock
   On Error Resume Next
   wskNetworkStatus(1).SendData "TEST"
   If Err <> 0 Or wskNetworkStatus(1).LocalIP = "127.0.0.1" Then
      txtStatus = "non siete connessi ad internet"
   Else
      txtStatus = "siete connessi ad internet"
   login.timer_stato_connessione2.Enabled = True
   End If
End Sub

Private Sub LoadNewWinsock()
   On Error Resume Next
   
   If wskNetworkStatus.Count > 1 Then
      If wskNetworkStatus(1).LocalIP <> "127.0.0.1" Then
         Exit Sub
      End If
   End If
   UnloadWinsock
   Load wskNetworkStatus(1)
   wskNetworkStatus(1).RemoteHost = wskNetworkStatus(1).LocalIP
   wskNetworkStatus(1).RemotePort = 5555
   wskNetworkStatus(1).LocalPort = 5555
   wskNetworkStatus(1).Protocol = sckUDPProtocol
   wskNetworkStatus(1).Bind
End Sub

Private Sub UnloadWinsock()
   If wskNetworkStatus.Count > 1 Then
      Unload wskNetworkStatus(1)
   End If
End Sub

