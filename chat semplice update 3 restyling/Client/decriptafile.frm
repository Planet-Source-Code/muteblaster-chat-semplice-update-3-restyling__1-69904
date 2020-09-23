VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form decriptafile 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "decripta file"
   ClientHeight    =   1470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   ScaleHeight     =   1470
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin client.CandyButton Cmdecrypt 
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   480
      Width           =   1695
      _ExtentX        =   2990
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
      Caption         =   "decripta"
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
   Begin client.CandyButton cmdBrowse 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
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
      Caption         =   "scegli file"
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
   Begin MSComDlg.CommonDialog Cd1 
      Left            =   1800
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "All  files (*.*)|*.*"
   End
   Begin VB.Shape Shape1 
      Height          =   1455
      Left            =   0
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "decriptafile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBrowse_Click()
On Error GoTo errhandle:
 Cd1.ShowOpen
 cmdCrypt.Enabled = True
 Exit Sub
errhandle:
 Exit Sub
End Sub

Private Sub decryptfile()
Dim iByte As Byte
 Dim iBytestr As String
 Dim iStr As String
 Dim i As Long
 
 Open Cd1.FileName For Binary As #1 '
 iStr = String(LOF(1), Chr(0))
 Get #1, , iStr

 For i = 1 To Len(iStr)
  iBytestr = Mid(iStr, i, 1)
  iByte = Asc(iBytestr)
  If iByte >= 0 And iByte <= 127 Then
   iByte = 128 + iByte
  Else
   iByte = iByte - 128
  End If
  Put #1, i, iByte
  DoEvents
 Next
 Close #1
End Sub

Private Sub Cmdecrypt_Click()
decryptfile
MsgBox " decryptaggio completato "
End Sub
