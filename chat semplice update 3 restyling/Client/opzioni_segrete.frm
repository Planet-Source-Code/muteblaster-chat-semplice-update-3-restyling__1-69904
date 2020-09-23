VERSION 5.00
Begin VB.Form opzioni_segrete 
   BackColor       =   &H8000000D&
   Caption         =   "opzioni segrete"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   10620
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      Caption         =   "sistema di crediti segreto"
      Height          =   4575
      Left            =   5520
      TabIndex        =   16
      Top             =   0
      Width           =   5055
   End
   Begin VB.Timer Timer_salva_selezione 
      Interval        =   5000
      Left            =   600
      Top             =   3840
   End
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin client.CandyButton Cmdsalva 
         Height          =   495
         Left            =   4320
         TabIndex        =   15
         Top             =   3960
         Width           =   855
         _ExtentX        =   1508
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
         Caption         =   "salva"
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
      Begin VB.CheckBox Check9 
         BackColor       =   &H80000009&
         Caption         =   "ATTIVA MODERAZIONE"
         Height          =   495
         Left            =   2640
         TabIndex        =   14
         Top             =   2880
         Width           =   2295
      End
      Begin VB.CheckBox Check8 
         BackColor       =   &H80000009&
         Caption         =   " MODERATORE IN PROVA"
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H8000000D&
         Height          =   1695
         Left            =   2640
         TabIndex        =   8
         Top             =   960
         Width           =   2295
         Begin VB.Frame Frame5 
            BackColor       =   &H8000000D&
            Height          =   975
            Left            =   240
            TabIndex        =   10
            Top             =   480
            Width           =   1815
            Begin VB.CheckBox Check7 
               BackColor       =   &H8000000D&
               Caption         =   "non magiorenne"
               Height          =   255
               Left            =   120
               TabIndex        =   12
               Top             =   600
               Width           =   1575
            End
            Begin VB.CheckBox Check6 
               BackColor       =   &H8000000D&
               Caption         =   "maggiorenne"
               Height          =   255
               Left            =   120
               TabIndex        =   11
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.CheckBox Check5 
            BackColor       =   &H8000000D&
            Caption         =   "seconda iscrizione"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   120
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H8000000E&
         Height          =   1695
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   2295
         Begin VB.CheckBox Check4 
            BackColor       =   &H80000009&
            Caption         =   "prima iscrizione"
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
            Left            =   240
            TabIndex        =   7
            Top             =   240
            Width           =   1695
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H80000009&
            Height          =   975
            Left            =   120
            TabIndex        =   4
            Top             =   600
            Width           =   2055
            Begin VB.CheckBox Check2 
               BackColor       =   &H80000009&
               Caption         =   "maggiorenne"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   6
               Top             =   240
               Width           =   1455
            End
            Begin VB.CheckBox Check3 
               BackColor       =   &H80000009&
               Caption         =   "non maggiorenne"
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
               Left            =   120
               TabIndex        =   5
               Top             =   600
               Width           =   1815
            End
         End
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000009&
         Caption         =   "blocca/sblocca keyboard e mouse"
         Height          =   495
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   1935
      End
      Begin client.chameleonButton chameleonButton1 
         Height          =   4575
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   8070
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
         MICON           =   "opzioni_segrete.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
End
Attribute VB_Name = "opzioni_segrete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Private Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)

Private Sub SaveControlValues()
 Call RegSave(Check1, Check1.Value)
 Call RegSave(Check2, Check2.Value)
 Call RegSave(Check3, Check3.Value)
 Call RegSave(Check4, Check4.Value)
 Call RegSave(Check5, Check5.Value)
 Call RegSave(Check6, Check6.Value)
 Call RegSave(Check7, Check7.Value)
 Call RegSave(Check8, Check8.Value)
 Call RegSave(Check9, Check9.Value)
End Sub

Private Sub Cmdsalva_Click()
 Call SaveControlValues
End Sub

Private Sub form_load()
 Check1.Value = RegLoad(Check1)
 Check2.Value = RegLoad(Check2)
 Check3.Value = RegLoad(Check3)
 Check4.Value = RegLoad(Check4)
 Check5.Value = RegLoad(Check5)
 Check6.Value = RegLoad(Check6)
 Check7.Value = RegLoad(Check7)
 Check8.Value = RegLoad(Check8)
 Check9.Value = RegLoad(Check9)
End Sub

Private Sub Timer_salva_selezione_Timer()
 Cmdsalva_Click
End Sub
