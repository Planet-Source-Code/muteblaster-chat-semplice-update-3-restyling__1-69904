VERSION 5.00
Begin VB.Form cercautente 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "cercautente"
   ClientHeight    =   2190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin client.CandyButton CmdSearch 
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1560
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
      Caption         =   "cerca"
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
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   2490
   End
   Begin client.CandyButton Cmdchiudi 
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
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
   Begin VB.Shape Shape1 
      Height          =   2175
      Left            =   0
      Top             =   0
      Width           =   3855
   End
   Begin VB.Image Image9 
      Height          =   450
      Left            =   0
      Picture         =   "cercautente.frx":0000
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image11 
      Height          =   435
      Left            =   3480
      Picture         =   "cercautente.frx":08B2
      Top             =   15
      Width           =   300
   End
   Begin VB.Image Image10 
      Height          =   450
      Left            =   240
      Picture         =   "cercautente.frx":0FC0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3285
   End
   Begin VB.Label Labelcerca 
      BackStyle       =   0  'Transparent
      Caption         =   "cerca utente per nick o per ip"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
End
Attribute VB_Name = "cercautente"
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

Private Sub Cmdchiudi_Click()
 chat.Picture8.Top = 10800
 Unload cercautente
End Sub

Private Sub CmdSearch_Click()
On Error GoTo CmdSearch_Click_Error
Dim i As Long, i2 As Long
Dim tmpStr As String

    tmpStr = Trim(txtSearch.Text)
    
    i2 = chat.listusers.ListCount - 1
    
    'In case of Listcount bug
    If i2 < -1 Then
        i2 = 40000
        On Error Resume Next
    End If
    
    For i = CLng(CmdSearch.Tag) + 1 To i2
        If InStr(1, UCase(chat.listusers.List(i)), UCase(tmpStr)) > 0 Then
            chat.listusers.Selected(i) = True
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
    
    With Me.txtSearch
        .SetFocus
        .SelStart = Len(.Text)
    End With

Exit Sub
CmdSearch_Click_Error:
End Sub

Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 ReleaseCapture
 SendMessage cercautente.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub

Private Sub txtSearch_Change()
On Error Resume Next

    With CmdSearch
        .Tag = -1
        .Caption = "Search"
    End With
    
    If Trim(txtSearch.Text) <> vbNullString And chat.listusers.ListCount <> 0 Then
        CmdSearch.Enabled = True
    Else
        CmdSearch.Enabled = False
    End If
 End Sub
 
 Private Sub FillListbox(lngItems As Long)
On Error Resume Next
Dim lngCount As Long, i As Long
    
    lngCount = chat.listusers.ListCount
    
    For i = (lngCount + 1) To (lngCount + lngItems)
        chat.listusers.AddItem "Item_" & i
    Next

End Sub
   

