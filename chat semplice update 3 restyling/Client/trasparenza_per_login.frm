VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form trasparenza_per_login 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   5280
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer_progressbar_per_trasparenza 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3840
      Top             =   6360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "200"
      Height          =   255
      Left            =   3960
      TabIndex        =   1
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "trasparenza_per_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Timer_progressbar_per_trasparenza_Timer()
 ProgressBar1.Value = ProgressBar1.Value + 1
 Label1.Caption = Label1.Caption - 24
 MakeTransparent Me.hwnd, Label1.Caption
 If ProgressBar1.Value = 100 Then
 Timer_progressbar_per_trasparenza.Enabled = False
 Unload trasparenza_per_login
 End If
End Sub
