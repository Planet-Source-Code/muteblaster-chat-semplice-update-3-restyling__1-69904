VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form caricamento 
   BorderStyle     =   0  'None
   ClientHeight    =   9510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   ScaleHeight     =   9510
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox frame5 
      Height          =   9495
      Left            =   0
      Picture         =   "caricamento.frx":0000
      ScaleHeight     =   9435
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.PictureBox Picture1 
         Height          =   375
         Left            =   360
         ScaleHeight     =   315
         ScaleWidth      =   3915
         TabIndex        =   8
         Top             =   1920
         Width           =   3975
         Begin VB.Label Label4 
            Caption         =   "il programma e' in fase di avvio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   0
            Width           =   3735
         End
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   285
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "1) caricamento opzioni di base completato"
         Top             =   5520
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "2) caricamento salvataggi completato"
         Top             =   6120
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "3) caricamento condizioni completato"
         Top             =   6720
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H8000000D&
         Height          =   735
         Left            =   360
         TabIndex        =   1
         Top             =   7440
         Width           =   3735
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   3360
            TabIndex        =   4
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   2880
            TabIndex        =   3
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Timer Timer_progressbar 
         Interval        =   12
         Left            =   3600
         Top             =   8280
      End
      Begin client.Anim Anim3 
         Height          =   2055
         Left            =   1200
         TabIndex        =   10
         Top             =   2760
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   3625
      End
      Begin VB.Image Image3 
         Height          =   1440
         Left            =   1560
         Picture         =   "caricamento.frx":13F46
         Top             =   120
         Width           =   1440
      End
   End
End
Attribute VB_Name = "caricamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer_progressbar_Timer()
  ProgressBar1.Value = ProgressBar1.Value + 1
  Label7.Caption = Label7.Caption + 1
  If ProgressBar1.Value = 100 Then
  Timer_progressbar.Enabled = False
  Label7.Caption = "100"
  End If
End Sub
