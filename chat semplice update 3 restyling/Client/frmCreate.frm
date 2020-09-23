VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmCreate 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "Crea Account"
   ClientHeight    =   8250
   ClientLeft      =   4455
   ClientTop       =   840
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   Begin client.CandyButton cmdCreate 
      Height          =   495
      Left            =   1800
      TabIndex        =   11
      Top             =   7680
      Width           =   1695
      _ExtentX        =   2990
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
      Caption         =   "crea account"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   6240
      Width           =   4335
      Begin VB.CommandButton Cmdprecedente 
         Height          =   255
         Left            =   3360
         TabIndex        =   22
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Cmdsucessivo 
         Height          =   255
         Left            =   3840
         TabIndex        =   21
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         Height          =   375
         Left            =   2040
         ScaleHeight     =   315
         ScaleWidth      =   1635
         TabIndex        =   20
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Txtprimaverifica 
         Height          =   285
         Left            =   2520
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Txtverificanumero 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3840
         TabIndex        =   18
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Txtverifica 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "scrivi il testo per completare la iscrizione"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.TextBox Txtserver 
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Top             =   5400
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock create 
      Left            =   240
      Top             =   7800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtVPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   4800
      Width           =   1815
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      MaxLength       =   12
      TabIndex        =   1
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Image Picture2 
      Height          =   735
      Left            =   120
      Picture         =   "frmCreate.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "egistrazione"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   1320
      TabIndex        =   26
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   960
      TabIndex        =   25
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "password predefinite"
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
      Left            =   1800
      TabIndex        =   24
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "aiuto a creare la password ?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1800
      TabIndex        =   23
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Shape Shape2 
      Height          =   1935
      Left            =   120
      Top             =   2400
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      Height          =   1095
      Left            =   120
      Top             =   960
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   3840
      TabIndex        =   17
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label_sicurezza3 
      BackColor       =   &H000000FF&
      Height          =   45
      Left            =   3000
      TabIndex        =   16
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label_sicurezza2 
      BackColor       =   &H000040C0&
      Height          =   45
      Left            =   2400
      TabIndex        =   15
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label_sicurezza1 
      BackColor       =   &H0000FF00&
      Height          =   45
      Left            =   1800
      TabIndex        =   14
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   2280
      TabIndex        =   13
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "massimo 20 caratteri"
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Labelserver 
      BackStyle       =   0  'Transparent
      Caption         =   "indirizzo dle server a cui connettersi"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Verifica Password: "
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   855
   End
End
Attribute VB_Name = "frmCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim verificanumero As Integer

Private Sub Cmdsucessivo_Click()
 If verificanumero < 3 Then verificanumero = verificanumero + 1
  Picture1.Picture = LoadPicture(App.Path & "\verifica iscrizione" & "\immagine" & verificanumero & ".jpg")
  Txtverificanumero.Text = verificanumero
End Sub

Private Sub Form_Load()
 cmdCreate.Enabled = False
 Cmdsucessivo_Click
End Sub

Private Sub cmdCreate_Click()
 
 If Not Txtverifica.Text = Txtprimaverifica.Text Then
    txtUsername.Text = ""
    Txtpassword.Text = ""
    txtVPassword.Text = ""
    Cmdsucessivo_Click
 End If
    If txtUsername <> "" Then
        If Txtpassword = txtVPassword Then
            create.Close
            create.connect txtServer.Text, 12584
        Else
            MsgBox "Password's don't match!", , "Error!"
        End If
    Else
        MsgBox "Enter a username!"
    End If
End Sub

Private Sub Cycle(ByRef verifica As image)
On Error GoTo ErrHandler
    verificanumero = verificanumero + 1
    Txtverificanumero.Text = verificanumero
    Picture1.Picture = LoadPicture(App.Path & "\verifica iscrizione" & "\immagine" & verificanumero & ".jpg")
Exit Sub
ErrHandler:
    verificanumero = 1
    Txtverificanumero.Text = verificanumero
    Picture1.Picture = LoadPicture(App.Path & "\verifica iscrizione" & "\immagine" & verificanumero & ".jpg")
    Resume Next
End Sub

Private Sub create_Connect()
    create.SendData "cre-" & txtUsername & "-" & Txtpassword & "\-"
End Sub

Private Sub create_DataArrival(ByVal bytesTotal As Long)
    create.GetData Msg, vbString
    If Left(Msg, 4) = "msg-" Then
        ' per arginare la posssibilita di creare molti account creiamo un form apposito dove ostiamo'
        ' gli account registrati quando se ne ha 2 ci si ferma'
       If account.txtUsername = "" Then
          account.txtUsername.Text = txtUsername.Text
          account.Txtpassword.Text = Txtpassword.Text
          Call RegSave(account.txtUsername, account.txtUsername.Text) ' salviamo nel form account la registrazione'
          Call RegSave(account.Txtpassword, account.Txtpassword.Text)
       ElseIf Not account.txtUsername = "" Then
          account.txtUsername2.Text = txtUsername.Text
          account.txtPassword2.Text = Txtpassword.Text
          Call RegSave(account.txtUsername2, account.txtUsername2.Text) ' salviamo nel form account la registrazione'
          Call RegSave(account.txtPassword2, account.txtPassword2.Text)
       End If
        MsgBox Mid(Msg, 5), , "Create Account"
        If psw_account.Txtpsw_account.Text = "" Then
           psw_account.Show
        Else
           create.Close
        End If
    End If
   opzioni_segrete.Check4 = 1
   Unload regole_canale
End Sub

' stabiliamo una certasicurezza in base al numero di caratteri della password'
' piu' caratteri ha' piu0 sicura e', ogni volta che si cupera un numero di'
' caratteri appare un label colorato per dirci quanto sicura e' '

Private Sub label7_change()
 
 If Label7.Caption > 5 Then
    Label_sicurezza3.Visible = True
 Else
    Label_sicurezza3.Visible = False
 End If
 If Label7.Caption > 3 Then
    Label_sicurezza2.Visible = True
 Else
    Label_sicurezza2.Visible = False
    Label7.ForeColor = &H80000012
 End If
 If Label7 > 1 Then
    Label_sicurezza1.Visible = True
 Else
    Label_sicurezza1.Visible = False
 End If
End Sub

Private Sub Label8_Click()
 crea_password_frmcreate.Show 1
End Sub

Private Sub Label8_MouseMove(label As Integer, Shift As Integer, X As Single, Y As Single)
 Label8.FontUnderline = True
End Sub

Private Sub label9_click()
 password_predefinite_registrazione.Show 1
End Sub

Private Sub Label9_MouseMove(label As Integer, Shift As Integer, X As Single, Y As Single)
 Label9.FontUnderline = True
End Sub

Private Sub Txtpassword_Change()
 If Txtpassword.Text = "" Then
    cmdCreate.Enabled = False
 End If
 Label7.Caption = Len(Txtpassword.Text)
End Sub

Private Sub txtPassword_MouseMove(Text As Integer, Shift As Integer, X As Single, Y As Single)
 Shape2.Visible = True
End Sub

Private Sub txtUsername_Change()
 If txtUsername.Text = "" Then
    cmdCreate.Enabled = False
 End If
 Label6.Caption = Len(txtUsername.Text)
End Sub

Private Sub txtUsername_MouseMove(Text As Integer, Shift As Integer, X As Single, Y As Single)
 Shape1.Visible = True
End Sub

Private Sub form_mousemove(Form As Integer, Shift As Integer, X As Single, Y As Single)
 Shape1.Visible = False
 Shape2.Visible = False
 Label8.FontUnderline = False
 Label9.FontUnderline = False
End Sub

Private Sub Txtverifica_Change()
 If Txtverifica.Text = "" Then
    cmdCreate.Enabled = False
  Else
    cmdCreate.Enabled = True
 End If
 ' per eliminare la sensibilita' delle maiuscole, facciamo in modo che ogni carattere'
 ' se e' maiuscolo venga ridimensionato'
 ' e' l'unica soluzione che mi e' venuta in mente'
 If Txtverifica.Text = "V8NXN" Then
    Txtverifica.Text = "v8nxn"
 ElseIf Txtverifica.Text = "A1YSP" Then
    Txtverifica.Text = "a1ysp"
 End If
End Sub

Private Sub Txtverificanumero_Change()
 If Txtverificanumero.Text = "1" Then
    Txtprimaverifica.Text = "v8nxn"
 ElseIf Txtverificanumero.Text = "2" Then
    Txtprimaverifica.Text = "a1ysp"
 End If
End Sub

Private Sub txtVPassword_Change()
 If txtVPassword.Text = "" Then
    cmdCreate.Enabled = False
 End If
End Sub
