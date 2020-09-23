VERSION 5.00
Begin VB.Form frmBuddyList 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   ClientHeight    =   9510
   ClientLeft      =   6030
   ClientTop       =   960
   ClientWidth     =   4740
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "fmBuddyList.frx":0000
   ScaleHeight     =   9510
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer_chiusura_winsock 
      Enabled         =   0   'False
      Interval        =   2300
      Left            =   4200
      Top             =   7680
   End
   Begin VB.Timer Timer_Text1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   4200
      Top             =   8040
   End
   Begin VB.Timer Timer_Txtricerca 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   4200
      Top             =   8520
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   4200
      Top             =   9000
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   3840
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   20
      Top             =   9120
      Visible         =   0   'False
      Width           =   375
      Begin VB.Timer Timer_browser 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   2760
         Top             =   120
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   18
      Text            =   "cerca sul web"
      Top             =   9000
      Width           =   2295
   End
   Begin VB.CommandButton CmdSearch 
      Caption         =   "search"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   7800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Txtricerca 
      Height          =   285
      Left            =   840
      TabIndex        =   16
      Text            =   "trova contatto"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Frame Frame12 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   840
      TabIndex        =   10
      Top             =   7800
      Width           =   3135
      Begin VB.TextBox Txtbannernumero 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Timer Timer_banner 
         Interval        =   10000
         Left            =   840
         Top             =   240
      End
      Begin VB.Image Image_banner 
         BorderStyle     =   1  'Fixed Single
         Height          =   975
         Left            =   0
         MousePointer    =   4  'Icon
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3135
      End
   End
   Begin client.CandyButton cmdChangeInfo 
      Height          =   375
      Left            =   0
      TabIndex        =   9
      ToolTipText     =   "cambia informazioni"
      Top             =   7920
      Visible         =   0   'False
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
      Caption         =   "cambia info"
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
   Begin client.CandyButton cmdDeleteBuddy 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "rimuovi contatto"
      Top             =   8520
      Visible         =   0   'False
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
      Caption         =   "rimuovi amico"
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
   Begin client.CandyButton cmdIM 
      Height          =   375
      Left            =   0
      TabIndex        =   7
      ToolTipText     =   "messaggi privati"
      Top             =   9000
      Visible         =   0   'False
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
      Caption         =   "PM"
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
   Begin client.CandyButton cmdGetInfo 
      Height          =   375
      Left            =   0
      TabIndex        =   6
      ToolTipText     =   "veedi informazioni"
      Top             =   8880
      Visible         =   0   'False
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
      Caption         =   "richiedi info"
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
   Begin client.CandyButton cmdAddBuddy 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      ToolTipText     =   "aggiungi contatto"
      Top             =   8160
      Visible         =   0   'False
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
      Caption         =   "aggiungi amico"
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
   Begin VB.ListBox lstOffline 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4530
      ItemData        =   "fmBuddyList.frx":13F46
      Left            =   2760
      List            =   "fmBuddyList.frx":13F48
      TabIndex        =   1
      Top             =   2640
      Width           =   1815
   End
   Begin VB.ListBox lstBuddy 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4530
      ItemData        =   "fmBuddyList.frx":13F4A
      Left            =   840
      List            =   "fmBuddyList.frx":13F4C
      TabIndex        =   0
      Top             =   2640
      Width           =   1815
   End
   Begin client.CandyButton Cmdexit 
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   3600
      Visible         =   0   'False
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
   Begin client.CandyButton Cmdritorna 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   7200
      Visible         =   0   'False
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
      Caption         =   "ritorna in chat"
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
   Begin client.CandyButton CandyButton2 
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   1800
      Width           =   375
      _ExtentX        =   661
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
      Caption         =   ""
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "fmBuddyList.frx":13F4E
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Image Image_ritorna_in_chat 
      Height          =   285
      Left            =   3360
      Picture         =   "fmBuddyList.frx":14C28
      Stretch         =   -1  'True
      ToolTipText     =   "ritorna in chat"
      Top             =   1200
      Width           =   735
   End
   Begin VB.Image Image_cambiainfo 
      Height          =   285
      Left            =   1920
      Picture         =   "fmBuddyList.frx":157D3
      Stretch         =   -1  'True
      ToolTipText     =   "cambia info"
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "banner gratuiti:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Line Line10 
      X1              =   120
      X2              =   720
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line9 
      X1              =   120
      X2              =   720
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line8 
      X1              =   120
      X2              =   720
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line7 
      X1              =   120
      X2              =   720
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line6 
      X1              =   120
      X2              =   720
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   720
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Image imgoffline 
      Height          =   255
      Index           =   0
      Left            =   2760
      Picture         =   "fmBuddyList.frx":15EBB
      Top             =   2400
      Width           =   240
   End
   Begin VB.Image imgonline 
      Height          =   240
      Index           =   1
      Left            =   840
      Picture         =   "fmBuddyList.frx":16279
      Top             =   2400
      Width           =   240
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
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
      Left            =   4320
      TabIndex        =   19
      Top             =   120
      Width           =   255
   End
   Begin VB.Image Image7 
      Height          =   495
      Left            =   220
      Picture         =   "fmBuddyList.frx":16603
      Stretch         =   -1  'True
      ToolTipText     =   "vai al sito di visual basic france"
      Top             =   6240
      Width           =   495
   End
   Begin VB.Image Image6 
      Height          =   495
      Left            =   220
      Picture         =   "fmBuddyList.frx":17BF0
      Stretch         =   -1  'True
      ToolTipText     =   "vai al sito di sourceforge"
      Top             =   5520
      Width           =   495
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   220
      Picture         =   "fmBuddyList.frx":188BA
      Stretch         =   -1  'True
      ToolTipText     =   "vai al sito gnu"
      Top             =   4800
      Width           =   495
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   220
      Picture         =   "fmBuddyList.frx":19584
      Stretch         =   -1  'True
      ToolTipText     =   "vai a goolge"
      Top             =   4080
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   220
      Picture         =   "fmBuddyList.frx":1A24E
      Stretch         =   -1  'True
      ToolTipText     =   "vai al sito di fuorissimo "
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   220
      Picture         =   "fmBuddyList.frx":1AF18
      Stretch         =   -1  'True
      ToolTipText     =   "vai a planet sourccode"
      Top             =   2640
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   285
      Left            =   3240
      Picture         =   "fmBuddyList.frx":1BBE2
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   345
   End
   Begin VB.Image image_ricerca 
      Height          =   300
      Left            =   2880
      Picture         =   "fmBuddyList.frx":1BEDF
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   300
   End
   Begin VB.Line Line5 
      X1              =   120
      X2              =   4680
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   120
      Y1              =   1680
      Y2              =   7440
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   4680
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4680
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Shape Shape4 
      Height          =   255
      Left            =   1440
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Shape Shape3 
      Height          =   255
      Left            =   1440
      Top             =   360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
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
      Left            =   1440
      TabIndex        =   14
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
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
      Left            =   1440
      TabIndex        =   13
      Top             =   360
      Width           =   2055
   End
   Begin VB.Image Picavatar 
      Height          =   1095
      Left            =   240
      Picture         =   "fmBuddyList.frx":1C1DC
      Stretch         =   -1  'True
      ToolTipText     =   "visualizza il tuo biglietto da visita"
      Top             =   240
      Width           =   975
   End
   Begin VB.Shape Shape2 
      Height          =   1455
      Left            =   120
      Top             =   120
      Width           =   4455
   End
   Begin VB.Image Image14 
      Height          =   285
      Left            =   2880
      Picture         =   "fmBuddyList.frx":1CEFA
      ToolTipText     =   "rimuovi utente"
      Top             =   1200
      Width           =   285
   End
   Begin VB.Image Image12 
      Height          =   285
      Left            =   2400
      Picture         =   "fmBuddyList.frx":1D3B0
      ToolTipText     =   "messaggi privati "
      Top             =   1200
      Width           =   285
   End
   Begin VB.Image Image13 
      Height          =   285
      Left            =   1560
      Picture         =   "fmBuddyList.frx":1D866
      Stretch         =   -1  'True
      ToolTipText     =   "richiedi info"
      Top             =   1200
      Width           =   285
   End
   Begin VB.Image Image15 
      Height          =   285
      Left            =   3360
      Picture         =   "fmBuddyList.frx":1DD1C
      ToolTipText     =   "aggiungi contatto"
      Top             =   1800
      Width           =   285
   End
   Begin VB.Shape Shape1 
      Height          =   9495
      Left            =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Offline:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Online:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   2400
      Width           =   735
   End
End
Attribute VB_Name = "frmBuddyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' GRAZIE A SANDRO FIRZZARIN, PROGRAMMATORE PROESSIONISTA PER IL GRANDE AIUTO'
' MELLA INTRODUZIONE DELLA MESSAGGISTICA ISTANTANEA'

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Dim bannernumero As Integer

Private Sub Cmd_fuorissimo_Click()

End Sub

Private Sub Cmd_planetsourcecode_Click()

End Sub

Private Sub CandyButton2_Click()
 Timer1.Enabled = True
End Sub

Private Sub Form_Load()
    For i = 0 To 50
        Set imwindow(i) = New frmIM
        DoEvents
    Next
   Label3.Caption = login.Txtnick.Text & "  " & login.Label8.Caption
   Label4.Caption = login.Txtfrase.Text
   Image_banner.Picture = LoadPicture(App.Path & "\immagini varie" & "\banner iniziale" & ".jpg")
End Sub


Private Sub cmdAddBuddy_Click()
    frmAdd.Show
End Sub

Private Sub cmdChangeInfo_Click()
    If LocalInfo <> "" Then
        frmSetInfo.txtInfo = Left(LocalInfo, Len(LocalInfo) - 2)
    End If
    frmSetInfo.txtInfo.SelStart = Len(frmSetInfo.txtInfo.Text)
    frmSetInfo.Show
End Sub

Private Sub cmdDeleteBuddy_Click()
    If lstBuddy.SelCount > 0 Then
        login.win.SendData "del-" & lstBuddy.List(lstBuddy.ListIndex) & "\-"
        
        lstBuddy.RemoveItem lstBuddy.ListIndex
    ElseIf lstOffline.SelCount > 0 Then
        login.win.SendData "del-" & lstOffline.List(lstOffline.ListIndex) & "\-"
        
        lstOffline.RemoveItem lstOffline.ListIndex
    End If
    
End Sub

Private Sub Cmdexit_Click()
frmBuddyList.Visible = False
End Sub

Private Sub cmdGetInfo_Click()
    If lstBuddy.SelCount > 0 Then
        login.win.SendData "ginf-" & lstBuddy.List(lstBuddy.ListIndex)
    ElseIf lstOffline.SelCount > 0 Then
        login.win.SendData "ginf-" & lstOffline.List(lstOffline.ListIndex)
    End If
End Sub

Private Sub cmdIM_Click()
    For i = 0 To 20
        If imwindow(i).Caption = "" Then
            imwindow(i).txtUser.Locked = False
            imwindow(i).Show
            i = 21
        End If
        DoEvents
    Next
End Sub

Private Sub Cmdritorna_Click()
 login.Timer_ritorna_in_chat1.Enabled = True
End Sub

Private Sub Command1_Click()

End Sub



Private Sub Form_Unload(Cancel As Integer)
    'End
End Sub

Private Sub Cycle(ByRef banner As image)
On Error GoTo ErrHandler
    bannernumero = bannernumero + 1
    Txtbannernumero.Text = bannernumero
    Image_banner.Picture = LoadPicture(App.Path & "\banner" & "\immagine" & bannernumero & ".jpg")
Exit Sub
ErrHandler:
    bannernumero = 1
    Txtbannernumero.Text = bannernumero
    Image_banner.Picture = LoadPicture(App.Path & "\banner" & "\immagine" & bannernumero & ".jpg")
    Resume Next
End Sub

Private Sub Image_banner_Click()
  webbrowser_banner.Timer1.Enabled = True
  webbrowser_banner.Show
 If Txtbannernumero.Text = 1 Then
    webbrowser_banner.browser.Navigate "http://80vogliadi.blogspot.com/"
 ElseIf Txtbannernumero.Text = 2 Then
    webbrowser_banner.browser.Navigate "http://www.planet-source-code.com/"
 ElseIf Txtbannernumero.Text = 3 Then
    webbrowser_banner.browser.Navigate "http://it.youtube.com/"
 End If
End Sub

Private Sub Image_banner_MouseMove(image As Integer, Shift As Integer, X As Single, Y As Single)
 Shape3.Visible = True
End Sub

Private Sub Image_cambiainfo_Click()
 cmdChangeInfo_Click
End Sub

Private Sub image_ricerca_Click()
 CmdSearch_Click
End Sub

Private Sub Image_ritorna_in_chat_Click()
 If chat.Visible = False Then
    Cmdritorna_Click
 End If
End Sub

Private Sub Image1_Click()
 webbrowser_banner.Timer1.Enabled = True
 webbrowser_banner.Show
 webbrowser_banner.browser.Navigate "http://www.google.it/search?q=" & Text1.Text
End Sub

Private Sub Image12_Click()
 cmdIM_Click
End Sub

Private Sub Image13_Click()
 cmdGetInfo_Click
End Sub

Private Sub Image14_Click()
 cmdDeleteBuddy_Click
End Sub

Private Sub Image15_Click()
 cmdAddBuddy_Click
End Sub

Private Sub Image2_Click()
 browser.browse.Navigate "http://www.planet-source-code.com/"
 Timer_browser.Enabled = True
End Sub

Private Sub Image3_Click()
 browser.browse.Navigate "http://www.fuorissimo.com/"
 Timer_browser.Enabled = True
End Sub

Private Sub Image4_Click()
 browser.browse.Navigate "http://www.google.it/"
 Timer_browser.Enabled = True
End Sub

Private Sub Image5_Click()
 browser.browse.Navigate "http://www.gnu.org/home.it.html"
 Timer_browser.Enabled = True
End Sub

Private Sub Image6_Click()
 browser.browse.Navigate "http://sourceforge.net/"
 Timer_browser.Enabled = True
End Sub

Private Sub Image7_Click()
 browser.browse.Navigate "http://www.vbfrance.com/"
 Timer_browser.Enabled = True
End Sub

Private Sub Label3_MouseMove(label As Integer, Shift As Integer, X As Single, Y As Single)
 Shape3.Visible = True
End Sub

Private Sub Label4_MouseMove(label As Integer, Shift As Integer, X As Single, Y As Single)
 Shape4.Visible = True
End Sub

Private Sub Label5_Click()
 'login.Timer_winsock_close.Enabled = True'
 If chat.Visible = True Then
   chat.Timeresci_chat.Enabled = True
   Timer_chiusura_winsock.Enabled = True
  ElseIf chat.Visible = False Then
    Timer_chiusura_winsock.Enabled = True
  End If
End Sub

Private Sub lstBuddy_DblClick()
    found = False
    countit = 0
    For i = 0 To 50
        If imwindow(i).Caption = "" And countit < 1 Then
            firstfree = i
            countit = 1
        End If
        If UCase(Left(imwindow(i).Caption, Len(lstBuddy.List(lstBuddy.ListIndex)))) = UCase(lstBuddy.List(lstBuddy.ListIndex)) Then
            imwindow(i).SetFocus
            found = True
            i = 51
        End If
        DoEvents
    Next
    If found = False Then
        imwindow(firstfree).Caption = lstBuddy.List(lstBuddy.ListIndex) & " : " & login.Txtnick
        imwindow(firstfree).txtUser = lstBuddy.List(lstBuddy.ListIndex)
        imwindow(firstfree).Show
        imwindow(firstfree).txtIm.SetFocus
    End If
End Sub

Private Sub lstBuddy_GotFocus()
    On Error Resume Next
    lstOffline.Selected(lstOffline.ListIndex) = False
End Sub

Private Sub lstOffline_GotFocus()
    On Error Resume Next
    lstBuddy.Selected(lstBuddy.ListIndex) = False
End Sub

Private Sub form_mousemove(Form As Integer, Shift As Integer, X As Single, Y As Single)
 Shape3.Visible = False
 Shape4.Visible = False
 Image2.Height = 495
 Image2.Width = 495
End Sub

 ' questo form e' borderless ( senza bordo), impostiamo la immagine 10'
' come bordo che gli permettera' di muovere il form'
Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage frmBuddyList.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub

Private Sub Picavatar_Click()
 biglietto_da_visita.Show 1
End Sub

Private Sub Text1_click()
 If Text1.Text = "cerca sul web" Then
    Text1.Text = ""
  If Txtricerca.Text = "" Then
    Txtricerca.Text = "trova contatto"
  End If
 End If
 Timer_Text1.Enabled = True
End Sub

Private Sub Timer_banner_Timer()
 Cycle Image_banner
End Sub

Private Sub CmdSearch_Click()
On Error GoTo CmdSearch_Click_Error
Dim i As Long, i2 As Long
Dim tmpStr As String

    tmpStr = Trim(Txtricerca.Text)
    
    i2 = lstBuddy.ListCount - 1
    
    'In case of Listcount bug
    If i2 < -1 Then
        i2 = 40000
        On Error Resume Next
    End If
    
    For i = CLng(CmdSearch.Tag) + 1 To i2
        If InStr(1, UCase(lstBuddy.List(i)), UCase(tmpStr)) > 0 Then
            lstBuddy.Selected(i) = True
            If i < i2 Then CmdSearch.Caption = "Next"
            CmdSearch.Tag = i
            If i = i2 Then Exit For
            Exit Sub
        End If
    Next
    
    With Me.CmdSearch
        .Caption = "Search"
        .Tag = -1
        .Enabled = False
    End With
    
    With Me.Txtricerca
        .SetFocus
        .SelStart = Len(.Text)
    End With

Exit Sub
CmdSearch_Click_Error:
End Sub
 
 Private Sub FillListbox(lngItems As Long)
On Error Resume Next
Dim lngCount As Long, i As Long
    
    lngCount = lstBuddy.ListCount
    
    For i = (lngCount + 1) To (lngCount + lngItems)
        lstBuddy.AddItem "Item_" & i
    Next

End Sub

Private Sub Timer_chiusura_winsock_Timer()
 login.Visible = True
   chat.Visible = False
   chat_style.Visible = False
   login.win.Close
   login.Ws.Close
   login.WsMSricevi.Close
   login.WsPMricevi.Close
   login.Wsricevifile.Close
   login.WsPMserverricevi.Close
   login.Wsricevicomandichat.Close
   login.Timer_unload_frmbuddylist.Enabled = True
   Timer_chiusura_winsock.Enabled = False
End Sub

Private Sub Timer_Text1_Timer()
 If Text1.Text = "" Then
    Text1.Text = "cerca sul web"
 End If
 Timer_Text1.Enabled = False
End Sub

Private Sub Timer_Txtricerca_Timer()
 If Txtricerca.Text = "" Then
    Txtricerca.Text = "trova contatto"
 End If
    Timer_Txtricerca.Enabled = False
End Sub

Private Sub Timer1_Timer()
 Picture1.Height = 255
 Picture1.Width = 375
 Picture1.Top = 9120
 Picture1.Left = 3840
 Picture1.Visible = False
 Timer1.Enabled = False
End Sub

Private Sub Timer_browser_Timer()
 SetParent browser.hwnd, Picture1.hwnd
 Picture1.Height = 4935
 Picture1.Width = 3735
 Picture1.Left = 840
 Picture1.Top = 2400
 browser.Show
 browser.Move 0, 0
 Picture1.Visible = True
 Timer_browser.Enabled = False
End Sub

Private Sub Timer2_Timer()

End Sub

Private Sub Txtricerca_Change()
 
 On Error Resume Next
    With CmdSearch
        .Tag = -1
        .Caption = "Search"
    End With
    
    If Trim(Txtricerca.Text) <> vbNullString And lstBuddy.ListCount <> 0 Then
        CmdSearch.Enabled = True
    Else
        CmdSearch.Enabled = False
    End If
End Sub

Private Sub Txtricerca_click()
 If Txtricerca.Text = "trova contatto" Then
    Txtricerca.Text = ""
  If Text1.Text = "" Then
    Text1.Text = "cerca sul web"
  End If
 End If
Timer_Txtricerca.Enabled = True
End Sub
