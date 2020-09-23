VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   BorderStyle     =   0  'None
   Caption         =   "Chat Example - Login"
   ClientHeight    =   3300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2025
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   2025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin client.CandyButton cmdCancel 
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   2640
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "cancella"
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
   Begin VB.Timer Timer_connessione 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   3000
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000000&
      Height          =   1575
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   1935
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   1680
         Top             =   0
      End
      Begin MSComctlLib.ImageList anim_browser 
         Left            =   0
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   22
         ImageHeight     =   22
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   17
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":062A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":0C54
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":127E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":18A8
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":1ED2
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":24FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":2B26
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":3150
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":377A
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":3DA4
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":43CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":49F8
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":5022
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":564C
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":5C76
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":62A0
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Image img_anim_browser 
         Height          =   330
         Left            =   720
         Picture         =   "frmClient.frx":68CA
         Top             =   600
         Width           =   330
      End
   End
   Begin MSWinsockLib.Winsock Ws 
      Left            =   120
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar Stat 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   7
      Top             =   3000
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Offline"
            TextSave        =   "Offline"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "8.47"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   3360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame fraLogin 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   1935
      Begin VB.CheckBox chkRemember 
         Appearance      =   0  'Flat
         Caption         =   "Remember Nickname and Server"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.TextBox txtServer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Text            =   "muteblaster.no-ip.info"
         Top             =   840
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtNick 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblServer 
         BackStyle       =   0  'Transparent
         Caption         =   "Server:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblNick 
         BackStyle       =   0  'Transparent
         Caption         =   "Nickname:"
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   120
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLogin_Click()
    'Attempt to Connect
    ChatUser = Txtnick.Text
    Call status("Connecting")
    Pause (1)
    Call SocketConnect("5106", WS)
End Sub
Private Sub Form_Load()
    Txtnick.Text = login.Txtnick.Text
    Timer_connessione.Enabled = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    Call CreateLog
    End
End Sub
' stabiliamo un timer che si occupera' della connessione'
Private Sub Timer_connessione_Timer()
  cmdLogin_Click
  Timer_connessione.Enabled = False
End Sub

Private Sub Timer1_Timer()
 Static frame As Integer
 frame = frame + 1
 If frame > anim_browser.ListImages.Count Then frame = 1
 img_anim_browser.Picture = anim_browser.ListImages(frame).Picture
End Sub

Private Sub Ws_Connect()
    'Socket Connected, Send [loginmulticanale] packet
    Call SendData(loginmulticanale(ChatUser))
    Call status("Connected")
    chat.Timer_setparent_frmrooms.Enabled = True
    chat.Picture22.Top = 4680
    chat.Picture22.Left = 120
End Sub
Private Sub Ws_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
    'Data is gotten.
    Call WS.GetData(Data$, vbString)
    'All data is seperated and checked in a function in the modData module.
    Call HandleData(Data$, WS)
End Sub
Private Sub Ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'Socket error, this will show if the [Server] isn't running
    Call status("Connection Error")
End Sub
