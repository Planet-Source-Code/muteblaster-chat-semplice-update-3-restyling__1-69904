VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmIM 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6765
   ClientLeft      =   5880
   ClientTop       =   4095
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtIm 
      Height          =   735
      Left            =   240
      TabIndex        =   7
      Top             =   5520
      Width           =   5655
   End
   Begin VB.CommandButton cmdGetInfo 
      Caption         =   "richiede Info"
      Height          =   375
      Left            =   6480
      TabIndex        =   4
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox txtHistory 
      Height          =   3255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   5741
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmIM.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add User"
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   3840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Default         =   -1  'True
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      Top             =   4800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Image Image_richiedi_info 
      Height          =   285
      Left            =   4200
      Picture         =   "frmIM.frx":0082
      ToolTipText     =   "cerca utente"
      Top             =   240
      Width           =   285
   End
   Begin VB.Image Image_addbuddy 
      Height          =   285
      Left            =   3720
      Picture         =   "frmIM.frx":0538
      ToolTipText     =   "aggiungi lista amici"
      Top             =   240
      Width           =   285
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "send"
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
      Left            =   6120
      TabIndex        =   8
      Top             =   5880
      Width           =   495
   End
   Begin VB.Image Image15 
      Height          =   6405
      Left            =   7560
      Picture         =   "frmIM.frx":09EE
      Stretch         =   -1  'True
      Top             =   360
      Width           =   75
   End
   Begin VB.Image Image14 
      Height          =   75
      Left            =   0
      Picture         =   "frmIM.frx":0A46
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   7935
   End
   Begin VB.Image Image21 
      Height          =   5820
      Left            =   0
      Picture         =   "frmIM.frx":0A9E
      Stretch         =   -1  'True
      Top             =   840
      Width           =   75
   End
   Begin VB.Image Image11 
      Height          =   675
      Left            =   6000
      Picture         =   "frmIM.frx":0AF6
      Top             =   5640
      Width           =   705
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00A06B4F&
      FillColor       =   &H80000014&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   240
      Top             =   5520
      Width           =   6570
   End
   Begin VB.Image Image2 
      Height          =   345
      Left            =   240
      Picture         =   "frmIM.frx":1A50
      Top             =   6360
      Width           =   90
   End
   Begin VB.Image Image3 
      Height          =   345
      Left            =   240
      Picture         =   "frmIM.frx":1B34
      Top             =   5160
      Width           =   90
   End
   Begin VB.Image Image7 
      Height          =   345
      Left            =   330
      Picture         =   "frmIM.frx":1C1B
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   6390
   End
   Begin VB.Image Image8 
      Height          =   345
      Left            =   6720
      Picture         =   "frmIM.frx":1D13
      Top             =   5160
      Width           =   90
   End
   Begin VB.Image Image9 
      Height          =   345
      Left            =   6720
      Picture         =   "frmIM.frx":1DFB
      Top             =   6360
      Width           =   90
   End
   Begin VB.Image Image10 
      Enabled         =   0   'False
      Height          =   345
      Left            =   330
      Picture         =   "frmIM.frx":1EDD
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   6390
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00814D3C&
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   7125
      Width           =   45
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   120
      Picture         =   "frmIM.frx":1FD5
      Top             =   960
      Width           =   90
   End
   Begin VB.Image Image5 
      Height          =   345
      Left            =   6600
      Picture         =   "frmIM.frx":202A
      Top             =   960
      Width           =   90
   End
   Begin VB.Image Image6 
      Height          =   345
      Left            =   210
      Picture         =   "frmIM.frx":207F
      Stretch         =   -1  'True
      Top             =   960
      Width           =   6390
   End
   Begin VB.Image Image16 
      Height          =   315
      Left            =   7440
      Picture         =   "frmIM.frx":20C9
      Top             =   0
      Width           =   150
   End
   Begin VB.Image Image17 
      Height          =   315
      Left            =   5880
      Picture         =   "frmIM.frx":232E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1575
   End
   Begin VB.Image Image18 
      Height          =   870
      Left            =   5080
      Picture         =   "frmIM.frx":23FA
      Top             =   0
      Width           =   855
   End
   Begin VB.Image Image19 
      Height          =   870
      Left            =   135
      Picture         =   "frmIM.frx":2AED
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4950
   End
   Begin VB.Image Image20 
      Height          =   870
      Left            =   0
      Picture         =   "frmIM.frx":2CAB
      Top             =   0
      Width           =   135
   End
   Begin VB.Image Image12 
      Height          =   4860
      Left            =   360
      Picture         =   "frmIM.frx":31DF
      Top             =   1920
      Width           =   7275
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.Image Image22 
      Height          =   9060
      Left            =   120
      Picture         =   "frmIM.frx":7FE1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8685
   End
End
Attribute VB_Name = "frmIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
    AddBuddy (txtUser)
End Sub

Private Sub cmdGetInfo_Click()
    If txtUser.Text <> "" Then
        login.win.SendData "ginf-" & txtUser.Text
    End If
End Sub

Private Sub cmdSend_Click()
     If txtIm <> "" And txtUser <> "" Then
        login.win.SendData "im-" & txtUser & "-" & txtIm & "\-"
        If txtHistory.Text = "" Then
            login.win.SendData "forsn-" & txtUser.Text & "\-"
        End If
        start = Len(txtHistory.Text)
        txtHistory.SelStart = Len(txtHistory.Text)
        txtHistory.SelText = login.Txtnick & ": " & txtIm.Text & vbCrLf
        txtHistory.SelStart = start
        txtHistory.SelLength = Len(login.Txtnick.Text)
        txtHistory.SelColor = vbBlue
        txtHistory.SelBold = True
        txtHistory.SelStart = Len(txtHistory)
        txtHistory.SelLength = 1
        If Me.Caption = "" Then
            Me.Caption = txtUser.Text & " : " & login.Txtnick.Text
        End If
        txtUser.Locked = True
        txtIm = ""
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Caption = ""
End Sub

Private Sub Image_addbuddy_Click()
 cmdAdd_Click
End Sub

Private Sub Image_richiedi_info_Click()
 cmdGetInfo_Click
End Sub

Private Sub Image11_Click()
 cmdSend_Click
End Sub

Private Sub Label2_Click()
 cmdSend_Click
End Sub
