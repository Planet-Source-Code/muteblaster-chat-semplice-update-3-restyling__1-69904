VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form MOD_chat 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "chat"
   ClientHeight    =   9285
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9285
   ScaleWidth      =   10740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   5055
      Left            =   240
      ScaleHeight     =   4995
      ScaleWidth      =   7755
      TabIndex        =   6
      Top             =   1800
      Width           =   7815
      Begin RichTextLib.RichTextBox txtchat 
         Height          =   4935
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   8705
         _Version        =   393217
         TextRTF         =   $"MOD_chat.frx":0000
      End
   End
   Begin client.CandyButton cmdsend 
      Height          =   615
      Left            =   7320
      TabIndex        =   5
      Top             =   8040
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1085
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
   Begin client.CandyButton Cmddisconnect 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
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
      Caption         =   "chiudi MOD chat"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   4
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Frame Frameusers 
      BackColor       =   &H80000013&
      Caption         =   "lista moderatori "
      Height          =   5415
      Left            =   8400
      TabIndex        =   1
      Top             =   1680
      Width           =   2175
      Begin VB.ListBox listusers 
         Height          =   4935
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.TextBox txtSend 
      Height          =   765
      Left            =   240
      TabIndex        =   0
      Top             =   7920
      Width           =   6915
   End
   Begin client.CandyButton Cmdexit 
      Height          =   255
      Left            =   9960
      TabIndex        =   3
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
   Begin VB.Shape Shape2 
      Height          =   9255
      Left            =   0
      Top             =   0
      Width           =   10695
   End
   Begin VB.Shape Shape1 
      Height          =   5295
      Left            =   120
      Top             =   1680
      Width           =   8055
   End
   Begin VB.Image txt1 
      Height          =   1530
      Left            =   240
      Picture         =   "MOD_chat.frx":0082
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   7800
   End
   Begin VB.Image snd1 
      Height          =   1545
      Left            =   120
      Picture         =   "MOD_chat.frx":4D44
      Top             =   7560
      Width           =   195
   End
   Begin VB.Image Image30 
      Height          =   1545
      Left            =   8040
      Picture         =   "MOD_chat.frx":5D9E
      Top             =   7560
      Width           =   195
   End
   Begin VB.Image Image5 
      Height          =   270
      Left            =   7200
      Picture         =   "MOD_chat.frx":6DF8
      Top             =   7620
      Width           =   285
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   120
      Picture         =   "MOD_chat.frx":7272
      Top             =   1200
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   450
      Left            =   360
      Picture         =   "MOD_chat.frx":7B24
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   7605
   End
   Begin VB.Image Image3 
      Height          =   435
      Left            =   7920
      Picture         =   "MOD_chat.frx":8ADE
      Top             =   1215
      Width           =   300
   End
   Begin VB.Image Image9 
      Height          =   450
      Left            =   0
      Picture         =   "MOD_chat.frx":91EC
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image11 
      Height          =   435
      Left            =   10320
      Picture         =   "MOD_chat.frx":9A9E
      Top             =   15
      Width           =   300
   End
   Begin VB.Image Image10 
      Height          =   450
      Left            =   240
      Picture         =   "MOD_chat.frx":A1AC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10125
   End
End
Attribute VB_Name = "MOD_chat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmdexit_Click()
 MOD_chat.Visible = False
End Sub

Private Sub cmdsend_Click()
If Not txtSend.Text = "" Then
    'txtchat.Text = txtchat.Text + " <" & MOD_login.Txtnick.Text & "> " & txtSend.Text + vbCrLf'
    MOD_login.WS.SendData " <" & MOD_login.Txtnick.Text & "> " & "<" & MOD_login.Txtfrase.Text & ">" & vbCrLf & txtSend.Text ' Sends the nick and the text in the send-box'
    txtSend.Text = "" ' dopo aver spedito il messaggio il txtsend ritorna vuoto'
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then ' se si preme enter'
    KeyAscii = 0
    If Not txtSend.Text = "" Then ' se il txtsend non e' vuoto'
        'txtMotta.Text = txtSend.Text + vbCrLf
        MOD_login.WS.SendData " <" & MOD_login.Txtnick.Text & "> " & "<" & MOD_login.Txtfrase.Text & ">" & vbCrLf & txtSend.Text ' Sends the nick and the text in the send-box'
        txtSend.Text = "" ' dopo aver spedito il messaggio il txtsend ritorna vuoto'
    End If
End If
End Sub

Private Sub Cmddisconnect_Click()
 MOD_login.Visible = True
 MOD_login.WS.SendData "@DISCONNECT:" & MOD_login.Txtnick.Text
 'Unload chat
End Sub


