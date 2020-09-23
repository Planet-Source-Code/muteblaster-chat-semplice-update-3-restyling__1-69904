VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form login 
   Caption         =   "chat semplice update 3"
   ClientHeight    =   9510
   ClientLeft      =   5115
   ClientTop       =   1215
   ClientWidth     =   4770
   ControlBox      =   0   'False
   Icon            =   "login.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "login.frx":038A
   ScaleHeight     =   9510
   ScaleWidth      =   4770
   Visible         =   0   'False
   Begin VB.Timer Timer_unload_frmbuddylist 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   3960
      Top             =   6000
   End
   Begin VB.Timer Timer_login_im 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3960
      Top             =   5520
   End
   Begin VB.Timer Timer_grafica_in_entrata 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3960
      Top             =   4800
   End
   Begin VB.Timer Timer_icona 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3960
      Top             =   4320
   End
   Begin VB.Timer Timer_ridimensionamento 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3960
      Top             =   3840
   End
   Begin VB.PictureBox Picture12 
      Height          =   9495
      Left            =   5760
      ScaleHeight     =   9435
      ScaleWidth      =   4755
      TabIndex        =   91
      Top             =   0
      Width           =   4815
   End
   Begin VB.TextBox Txtnick 
      BackColor       =   &H80000013&
      Height          =   285
      Left            =   1080
      TabIndex        =   64
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox txtPass 
      BackColor       =   &H80000013&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   63
      Top             =   3480
      Width           =   2655
   End
   Begin VB.CheckBox Checkremember 
      BackColor       =   &H80000013&
      Caption         =   "ricorda nick , password e frase tipica"
      Height          =   255
      Left            =   1080
      TabIndex        =   62
      Top             =   4920
      Width           =   255
   End
   Begin VB.CheckBox Checkautoconnect 
      BackColor       =   &H80000013&
      Caption         =   "connetti automaticamente all' avvio"
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
      Left            =   1080
      TabIndex        =   61
      Top             =   5280
      Width           =   255
   End
   Begin VB.CheckBox Checkattivasfondi_txtchat 
      BackColor       =   &H80000013&
      Caption         =   "attiva sfondi chat"
      Height          =   255
      Left            =   1080
      TabIndex        =   60
      Top             =   5640
      Width           =   255
   End
   Begin VB.TextBox Txtip 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      ForeColor       =   &H80000002&
      Height          =   285
      Left            =   1560
      TabIndex        =   59
      Text            =   "localhost"
      Top             =   7200
      Width           =   2655
   End
   Begin VB.Frame framelogin 
      BackColor       =   &H80000013&
      Caption         =   "login"
      Height          =   255
      Left            =   3840
      TabIndex        =   58
      Top             =   2520
      Width           =   615
      Begin MSComctlLib.ImageList status 
         Left            =   240
         Top             =   3240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "login.frx":142D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "login.frx":14622
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "login.frx":14944
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "login.frx":151BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "login.frx":15F5C
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "login.frx":16E22
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "login.frx":17CE8
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList anim_browser 
         Left            =   840
         Top             =   3240
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
               Picture         =   "login.frx":18B1A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "login.frx":19144
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "login.frx":1976E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "login.frx":19D98
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "login.frx":1A3C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "login.frx":1A9EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "login.frx":1B016
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "login.frx":1B640
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "login.frx":1BC6A
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "login.frx":1C294
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "login.frx":1C8BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "login.frx":1CEE8
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "login.frx":1D512
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "login.frx":1DB3C
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "login.frx":1E166
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "login.frx":1E790
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "login.frx":1EDBA
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H80000013&
      Height          =   6135
      Left            =   11280
      TabIndex        =   44
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
      Begin VB.Frame Frame11 
         BackColor       =   &H80000013&
         Caption         =   "Frame11"
         Height          =   3615
         Left            =   120
         TabIndex        =   47
         Top             =   1080
         Width           =   975
         Begin VB.Frame Frame_siti_amici 
            BackColor       =   &H80000013&
            BorderStyle     =   0  'None
            Height          =   5895
            Left            =   0
            TabIndex        =   48
            Top             =   -600
            Width           =   975
            Begin client.CandyButton Cmd_vbfrance 
               Height          =   495
               Left            =   120
               TabIndex        =   54
               ToolTipText     =   "vai al siti di visual basic france"
               Top             =   5160
               Width           =   735
               _ExtentX        =   1296
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
               Caption         =   ""
               IconHighLiteColor=   0
               CaptionHighLiteColor=   0
               Picture         =   "login.frx":1F3E4
               Style           =   1
               Checked         =   0   'False
               ColorButtonHover=   16760976
               ColorButtonUp   =   15309136
               ColorButtonDown =   15309136
               BorderBrightness=   0
               ColorBright     =   16772528
               DisplayHand     =   0   'False
               ColorScheme     =   0
            End
            Begin client.CandyButton Cmd_sourceforge 
               Height          =   495
               Left            =   120
               TabIndex        =   53
               ToolTipText     =   "vai a sourceforge"
               Top             =   4320
               Width           =   735
               _ExtentX        =   1296
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
               Caption         =   ""
               IconHighLiteColor=   0
               CaptionHighLiteColor=   0
               Picture         =   "login.frx":200BE
               Style           =   1
               Checked         =   0   'False
               ColorButtonHover=   16760976
               ColorButtonUp   =   15309136
               ColorButtonDown =   15309136
               BorderBrightness=   0
               ColorBright     =   16772528
               DisplayHand     =   0   'False
               ColorScheme     =   0
            End
            Begin client.CandyButton Cmd_gnu 
               Height          =   495
               Left            =   120
               TabIndex        =   52
               Top             =   3480
               Width           =   735
               _ExtentX        =   1296
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
               Caption         =   ""
               IconHighLiteColor=   0
               CaptionHighLiteColor=   0
               Picture         =   "login.frx":20D98
               Style           =   1
               Checked         =   0   'False
               ColorButtonHover=   16760976
               ColorButtonUp   =   15309136
               ColorButtonDown =   15309136
               BorderBrightness=   0
               ColorBright     =   16772528
               DisplayHand     =   0   'False
               ColorScheme     =   0
            End
            Begin client.CandyButton Cmdgoogle 
               Height          =   615
               Left            =   120
               TabIndex        =   49
               ToolTipText     =   "cerca con google"
               Top             =   2640
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
               Caption         =   ""
               IconHighLiteColor=   0
               CaptionHighLiteColor=   0
               Picture         =   "login.frx":21A72
               Style           =   1
               Checked         =   0   'False
               ColorButtonHover=   16760976
               ColorButtonUp   =   15309136
               ColorButtonDown =   15309136
               BorderBrightness=   0
               ColorBright     =   16772528
               DisplayHand     =   0   'False
               ColorScheme     =   0
            End
            Begin client.CandyButton Cmd_fuorissimo 
               Height          =   615
               Left            =   120
               TabIndex        =   50
               ToolTipText     =   "vai al sito di fuorissimo"
               Top             =   1800
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
               Caption         =   ""
               IconHighLiteColor=   0
               CaptionHighLiteColor=   0
               Picture         =   "login.frx":2274C
               Style           =   1
               Checked         =   0   'False
               ColorButtonHover=   16760976
               ColorButtonUp   =   15309136
               ColorButtonDown =   15309136
               BorderBrightness=   0
               ColorBright     =   16772528
               DisplayHand     =   0   'False
               ColorScheme     =   0
            End
            Begin client.CandyButton Cmd_planetsourcecode 
               Height          =   615
               Left            =   120
               TabIndex        =   51
               ToolTipText     =   "vai a planet source code"
               Top             =   960
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
               Caption         =   ""
               IconHighLiteColor=   0
               CaptionHighLiteColor=   0
               Picture         =   "login.frx":23426
               Style           =   1
               Checked         =   0   'False
               ColorButtonHover=   16760976
               ColorButtonUp   =   15309136
               ColorButtonDown =   15309136
               BorderBrightness=   0
               ColorBright     =   16772528
               DisplayHand     =   0   'False
               ColorScheme     =   0
            End
            Begin VB.Line Line9 
               X1              =   120
               X2              =   840
               Y1              =   5760
               Y2              =   5760
            End
            Begin VB.Line Line8 
               X1              =   120
               X2              =   840
               Y1              =   5040
               Y2              =   5040
            End
            Begin VB.Line Line7 
               X1              =   0
               X2              =   840
               Y1              =   3360
               Y2              =   3360
            End
            Begin VB.Line Line6 
               X1              =   0
               X2              =   840
               Y1              =   2520
               Y2              =   2520
            End
            Begin VB.Line Line4 
               X1              =   0
               X2              =   840
               Y1              =   4200
               Y2              =   4200
            End
            Begin VB.Line Line3 
               X1              =   960
               X2              =   960
               Y1              =   840
               Y2              =   5760
            End
            Begin VB.Line Line2 
               X1              =   840
               X2              =   0
               Y1              =   1680
               Y2              =   1680
            End
            Begin VB.Line Line1 
               X1              =   0
               X2              =   960
               Y1              =   840
               Y2              =   840
            End
         End
      End
      Begin client.CandyButton Cmd_scorrere_in_su 
         Height          =   255
         Left            =   360
         TabIndex        =   45
         Top             =   4800
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "^"
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
      Begin client.CandyButton Cmdscorrere_in_giu 
         Height          =   255
         Left            =   360
         TabIndex        =   46
         Top             =   5040
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
         Caption         =   ""
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
         Height          =   615
         Left            =   240
         TabIndex        =   55
         Top             =   360
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
         Caption         =   ""
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Picture         =   "login.frx":24100
         Style           =   1
         Checked         =   0   'False
         ColorButtonHover=   16760976
         ColorButtonUp   =   15309136
         ColorButtonDown =   15309136
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   0
      End
   End
   Begin VB.PictureBox Picture11 
      Height          =   8295
      Left            =   240
      ScaleHeight     =   8235
      ScaleWidth      =   6315
      TabIndex        =   43
      Top             =   12480
      Width           =   6375
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000D&
      Height          =   2535
      Left            =   240
      TabIndex        =   40
      Top             =   11280
      Width           =   2775
      Begin client.CandyButton Cchiudi 
         Height          =   255
         Left            =   2040
         TabIndex        =   41
         Top             =   120
         Width           =   495
         _ExtentX        =   873
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
         Caption         =   "x"
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
      Begin client.CandyButton Cmdbiglietto 
         Height          =   375
         Left            =   240
         TabIndex        =   42
         Top             =   840
         Width           =   2055
         _ExtentX        =   3625
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
         Caption         =   "crea biglietto da visita"
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
   End
   Begin VB.PictureBox Picture10 
      Height          =   8415
      Left            =   240
      ScaleHeight     =   8355
      ScaleWidth      =   6315
      TabIndex        =   39
      Top             =   11280
      Width           =   6375
   End
   Begin VB.PictureBox Picture9 
      BackColor       =   &H80000013&
      Height          =   9135
      Left            =   0
      ScaleHeight     =   9075
      ScaleWidth      =   6915
      TabIndex        =   29
      Top             =   11280
      Width           =   6975
      Begin VB.Frame Frame9 
         BackColor       =   &H80000013&
         Height          =   4095
         Left            =   1200
         TabIndex        =   32
         Top             =   3360
         Width           =   3975
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "6) chiusura sesta parte efettuata"
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
            TabIndex        =   38
            Top             =   3600
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "5) chiusura quinta parte effettuata"
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
            TabIndex        =   37
            Top             =   3120
            Visible         =   0   'False
            Width           =   3015
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "4) chiusura quarta parte effettuata"
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
            TabIndex        =   36
            Top             =   2520
            Visible         =   0   'False
            Width           =   3015
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "3) chiusura terza parte effettuata"
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
            TabIndex        =   35
            Top             =   1920
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "2) chiusura seconda parte effettuata"
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
            TabIndex        =   34
            Top             =   1320
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "1) chiusura prima parte effettuata"
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
            TabIndex        =   33
            Top             =   720
            Visible         =   0   'False
            Width           =   2895
         End
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "e' incorso la chiusura del programma"
         Top             =   2520
         Width           =   4575
      End
      Begin client.chameleonButton chameleonButton9 
         Height          =   9015
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   15901
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
         MICON           =   "login.frx":24DDA
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
   Begin VB.PictureBox Picture8 
      BackColor       =   &H80000013&
      Height          =   6375
      Left            =   240
      ScaleHeight     =   6315
      ScaleWidth      =   4755
      TabIndex        =   28
      Top             =   11280
      Width           =   4815
   End
   Begin VB.PictureBox Picture7 
      Height          =   4575
      Left            =   240
      ScaleHeight     =   4515
      ScaleWidth      =   6315
      TabIndex        =   27
      Top             =   11280
      Width           =   6375
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H80000013&
      Height          =   8295
      Left            =   240
      ScaleHeight     =   8235
      ScaleWidth      =   6315
      TabIndex        =   26
      Top             =   11280
      Width           =   6375
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H80000013&
      Height          =   2415
      Left            =   240
      ScaleHeight     =   2355
      ScaleWidth      =   2595
      TabIndex        =   25
      Top             =   11280
      Width           =   2655
   End
   Begin VB.PictureBox Picture4 
      Height          =   2775
      Left            =   6720
      ScaleHeight     =   2715
      ScaleWidth      =   4395
      TabIndex        =   24
      Top             =   11280
      Width           =   4455
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H80000013&
      Height          =   8415
      Left            =   240
      TabIndex        =   20
      Top             =   12000
      Width           =   6375
      Begin VB.TextBox Text8 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   615
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "sei in chat"
         Top             =   3720
         Width           =   2055
      End
      Begin client.chameleonButton chameleonButton8 
         Height          =   8415
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   14843
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
         MICON           =   "login.frx":24DF6
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
   Begin VB.PictureBox Picture2 
      Height          =   8295
      Left            =   18000
      ScaleHeight     =   8235
      ScaleWidth      =   6315
      TabIndex        =   18
      Top             =   360
      Width           =   6375
      Begin client.chameleonButton chameleonButton7 
         Height          =   8295
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   14631
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
         MICON           =   "login.frx":24E12
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
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Height          =   8175
      Left            =   240
      TabIndex        =   14
      Top             =   11280
      Width           =   6375
      Begin VB.PictureBox Picture3 
         Height          =   3375
         Left            =   1560
         ScaleHeight     =   3315
         ScaleWidth      =   3315
         TabIndex        =   23
         Top             =   4080
         Width           =   3375
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   420
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "accessoalla chat in corso"
         Top             =   840
         Width           =   3375
      End
      Begin client.Anim Anim2 
         Height          =   1575
         Left            =   2160
         TabIndex        =   15
         Top             =   1920
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   2778
      End
      Begin client.chameleonButton chameleonButton6 
         Height          =   2295
         Left            =   1800
         TabIndex        =   16
         Top             =   1560
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   4048
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
         MICON           =   "login.frx":24E2E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Image img_back 
         Height          =   8175
         Left            =   0
         Picture         =   "login.frx":24E4A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   6375
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000013&
      Height          =   4455
      Left            =   2880
      TabIndex        =   13
      Top             =   9720
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Timer Timer_ridimensionamento_in_grandezza 
         Enabled         =   0   'False
         Interval        =   32
         Left            =   2760
         Top             =   3120
      End
      Begin VB.Timer timer_stato_connessione2 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   2280
         Top             =   3120
      End
      Begin VB.Timer Timer_tempo_di_connessione 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   840
         Top             =   2640
      End
      Begin VB.Timer Timer_stato_connessione 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   1320
         Top             =   3120
      End
      Begin VB.Timer Timer_unload7 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   2160
         Top             =   840
      End
      Begin VB.Timer Timer_unload6 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   1680
         Top             =   1080
      End
      Begin VB.Timer Timer_unload5 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   1200
         Top             =   1080
      End
      Begin VB.Timer Timer_unload4 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   720
         Top             =   1080
      End
      Begin VB.Timer Timer_unload3 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   1680
         Top             =   600
      End
      Begin VB.Timer Timer_unload2 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   1200
         Top             =   600
      End
      Begin VB.Timer Timer_unload1 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   720
         Top             =   600
      End
      Begin VB.Timer Timer_bannaggio 
         Enabled         =   0   'False
         Interval        =   4000
         Left            =   1800
         Top             =   2160
      End
      Begin VB.Timer Timer_lmage1 
         Interval        =   180
         Left            =   840
         Top             =   2160
      End
      Begin VB.Timer Timer_attesa 
         Enabled         =   0   'False
         Interval        =   3000
         Left            =   360
         Top             =   2160
      End
      Begin VB.Timer Timer_ritardo 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   2280
         Top             =   2160
      End
      Begin VB.Timer Timer_recconnect 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   1320
         Top             =   2160
      End
      Begin VB.Timer Timer_caricamento_componenti_allavvio 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   360
         Top             =   2640
      End
      Begin VB.Timer Timer_check 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   1320
         Top             =   2640
      End
      Begin VB.Timer Timer_richiamo_salvataggio_opzioni_segrete 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   2760
         Top             =   2160
      End
      Begin VB.Timer Timer_winsock_listen 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   1800
         Top             =   2640
      End
      Begin VB.Timer Timer_winsock_close 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   2760
         Top             =   2640
      End
      Begin VB.Timer Timer_login_chat1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   360
         Top             =   3120
      End
      Begin VB.Timer Timer_login_chat2 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   1800
         Top             =   3120
      End
      Begin VB.Timer Timer_ritorna_in_chat1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   2280
         Top             =   2640
      End
      Begin VB.Timer Timer_ritorna_in_chat2 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   840
         Top             =   3120
      End
      Begin MSWinsockLib.Winsock WS 
         Left            =   3960
         Top             =   840
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock WsPMricevi 
         Left            =   4560
         Top             =   840
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   3333
      End
      Begin MSWinsockLib.Winsock WsMSricevi 
         Left            =   3960
         Top             =   1320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   3334
      End
      Begin MSWinsockLib.Winsock Wsricevifile 
         Left            =   4560
         Top             =   1320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   3335
      End
      Begin MSComctlLib.ImageList anim 
         Left            =   5640
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   56
         ImageHeight     =   39
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "login.frx":AE190
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "login.frx":AFCBA
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "login.frx":B17E4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "login.frx":B330E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSWinsockLib.Winsock WsPMserverricevi 
         Left            =   3960
         Top             =   1800
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   3336
      End
      Begin MSWinsockLib.Winsock win 
         Left            =   4560
         Top             =   1800
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock Wsricevicomandichat 
         Left            =   3960
         Top             =   2280
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   3337
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "timer svolgimento programma"
         Height          =   255
         Left            =   360
         TabIndex        =   57
         Top             =   1800
         Width           =   2775
      End
      Begin VB.Shape Shape4 
         Height          =   2055
         Left            =   240
         Top             =   1800
         Width           =   3135
      End
      Begin VB.Shape Shape3 
         Height          =   1455
         Left            =   360
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000013&
         Caption         =   "timer per la chiusura del programma"
         Height          =   1335
         Left            =   480
         TabIndex        =   56
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H80000013&
      Height          =   8295
      Left            =   18000
      TabIndex        =   8
      Top             =   9120
      Width           =   6375
      Begin VB.Timer Timer_psw_sicurezza_sbagliata 
         Enabled         =   0   'False
         Interval        =   30000
         Left            =   720
         Top             =   5520
      End
      Begin client.CandyButton Cmdchiudiprogramma 
         Height          =   495
         Left            =   3720
         TabIndex        =   12
         Top             =   6120
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "chiudi programma"
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
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   1800
         TabIndex        =   11
         Text            =   "WARNING"
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   10
         Text            =   "login.frx":B4E38
         Top             =   1680
         Width           =   5655
      End
      Begin client.chameleonButton chameleonButton5 
         Height          =   8295
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   14631
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
         MICON           =   "login.frx":B4EE7
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
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000D&
      Caption         =   " password di accesso"
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   11280
      Width           =   6375
      Begin VB.TextBox Text4 
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2160
         Width           =   1335
      End
      Begin client.CandyButton Cmdrecuper 
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2400
         Width           =   2775
         _ExtentX        =   4895
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
         Caption         =   "recupero password dimenticata"
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
      Begin client.CandyButton Cmdverifica 
         Height          =   375
         Left            =   3240
         TabIndex        =   4
         Top             =   2040
         Width           =   1095
         _ExtentX        =   1931
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
         Caption         =   "verifica"
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
      Begin VB.TextBox Txtpassword_sicurezza 
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   3975
      End
      Begin client.chameleonButton chameleonButton3 
         Height          =   855
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   1508
         BTYPE           =   3
         TX              =   "chameleonButton3"
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
         MICON           =   "login.frx":B4F03
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Shape Shape2 
         Height          =   975
         Left            =   4440
         Shape           =   4  'Rounded Rectangle
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "massimo 4 tentativi"
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
         Left            =   4560
         TabIndex        =   7
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "scrivi la password di sicurezza che hai impostato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   4095
      End
   End
   Begin client.CandyButton Cmddisconnetti 
      Height          =   375
      Left            =   3240
      TabIndex        =   65
      Top             =   1320
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
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
      Caption         =   "disconnetti"
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
   Begin client.CandyButton CandyButton_avatar 
      Height          =   375
      Left            =   3240
      TabIndex        =   66
      Top             =   1920
      Width           =   1095
      _ExtentX        =   1931
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
      Caption         =   "scegli avatar"
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
   Begin RichTextLib.RichTextBox Txtfrase 
      Height          =   375
      Left            =   1080
      TabIndex        =   67
      Top             =   4080
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393217
      BackColor       =   -2147483629
      Enabled         =   -1  'True
      TextRTF         =   $"login.frx":B4F1F
   End
   Begin client.CandyButton Cmdannulla 
      Height          =   375
      Left            =   1440
      TabIndex        =   68
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
      Caption         =   "annulla"
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
   Begin client.CandyButton cmdLogin 
      Height          =   375
      Left            =   1920
      TabIndex        =   69
      Top             =   6000
      Width           =   855
      _ExtentX        =   1508
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
      Caption         =   "accedi"
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
      TabIndex        =   70
      Top             =   9000
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
   Begin client.Anim Anim1 
      Height          =   1215
      Left            =   1560
      TabIndex        =   71
      Top             =   7800
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2143
   End
   Begin client.Anim Anim4 
      Height          =   735
      Left            =   3480
      TabIndex        =   72
      Top             =   8520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
   End
   Begin VB.Image pointer_icon 
      Height          =   480
      Left            =   3000
      Picture         =   "login.frx":B4FA1
      Top             =   11000
      Width           =   480
   End
   Begin VB.Image Image_angolo 
      Height          =   660
      Left            =   4050
      Picture         =   "login.frx":B50F3
      Top             =   8800
      Width           =   660
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H8000000D&
      FillColor       =   &H00FF0000&
      Height          =   210
      Left            =   2880
      Top             =   195
      Width           =   255
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "^"
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
      Left            =   2950
      TabIndex        =   97
      ToolTipText     =   "opzioni"
      Top             =   240
      Width           =   255
   End
   Begin VB.Line Line5 
      X1              =   3240
      X2              =   3240
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4080
      TabIndex        =   96
      Top             =   30
      Width           =   255
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H8000000D&
      FillColor       =   &H8000000D&
      Height          =   210
      Left            =   4080
      Top             =   195
      Width           =   240
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3720
      TabIndex        =   95
      ToolTipText     =   "ridimensiona"
      Top             =   150
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H80000002&
      Height          =   210
      Left            =   3360
      Top             =   195
      Width           =   240
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   94
      Top             =   4635
      Width           =   255
   End
   Begin VB.Shape Shape5 
      Height          =   255
      Left            =   1680
      Top             =   4560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1800
      TabIndex        =   93
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "stato :"
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
      Left            =   1080
      TabIndex        =   92
      Top             =   4560
      Width           =   615
   End
   Begin VB.Image Picavatar 
      Height          =   1455
      Left            =   1680
      Picture         =   "login.frx":B694D
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   0
      Picture         =   "login.frx":B766B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   645
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "LOGIN :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   840
      TabIndex        =   90
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label_tempo_di_connessione 
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
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   1680
      TabIndex        =   89
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Cmdminimizza 
      BackStyle       =   0  'Transparent
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   3360
      TabIndex        =   88
      ToolTipText     =   "minimizza"
      Top             =   150
      Width           =   255
   End
   Begin VB.Label Cmdexit 
      BackStyle       =   0  'Transparent
      Caption         =   "x"
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
      Height          =   255
      Left            =   4440
      TabIndex        =   87
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Labelnick 
      BackStyle       =   0  'Transparent
      Caption         =   "login"
      Height          =   255
      Left            =   1080
      TabIndex        =   86
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   255
      Left            =   1080
      TabIndex        =   85
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Labelfrase 
      BackStyle       =   0  'Transparent
      Caption         =   "frase tipica"
      Height          =   255
      Left            =   1080
      TabIndex        =   84
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "ricorda nick , password e frase tipica"
      Height          =   255
      Left            =   1440
      TabIndex        =   83
      Top             =   4920
      Width           =   3255
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "connetti automaticamente all' avvio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   1440
      TabIndex        =   82
      Top             =   5280
      Width           =   3135
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "attiva sfondi chat"
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   1440
      TabIndex        =   81
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Label Labelmioip 
      BackStyle       =   0  'Transparent
      Caption         =   "mio ip"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   80
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1200
      TabIndex        =   79
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Labelip 
      BackStyle       =   0  'Transparent
      Caption         =   "indirizzo del server"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   78
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Label Cmdprofilo 
      BackStyle       =   0  'Transparent
      Caption         =   "crea profilo"
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
      Left            =   120
      TabIndex        =   77
      Top             =   7800
      Width           =   1335
   End
   Begin VB.Label cmdCreate 
      BackStyle       =   0  'Transparent
      Caption         =   "crea account"
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
      Left            =   120
      TabIndex        =   76
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Label Label_vediaccount 
      BackStyle       =   0  'Transparent
      Caption         =   "vedi account"
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
      Left            =   120
      TabIndex        =   75
      Top             =   8520
      Width           =   1695
   End
   Begin VB.Label Cmdrouter 
      BackStyle       =   0  'Transparent
      Caption         =   "router"
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
      Left            =   3480
      TabIndex        =   74
      Top             =   7800
      Width           =   855
   End
   Begin VB.Label Cmdtest 
      BackStyle       =   0  'Transparent
      Caption         =   "testa porte"
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
      Left            =   3480
      TabIndex        =   73
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      Height          =   9495
      Left            =   0
      Top             =   0
      Width           =   4740
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim frmResize As New ControlResizer
Dim Newalert As New cpopup ' dichiariamo la variabile che mi permettera
                           ' di aprire nuovi popup ogni volta '
                           ' richiamando il modulo di classe cpopup'
Dim BlnTflag As Boolean ' Transfert Flag'
Dim LngCursor As Long ' source file position pointer'
Dim tentativonumero As Integer ' il numero di tentativi possibili cin la psw di sicurezza'
Dim bannernumero As Integer
Dim StartTime ' dichiariamo la variabile che mi permettera' di contare il tempo di connessione , per un futuro sistema di crediti'

' allavvio del form load ( cioe' del programma ) i vari winsock che gestiscono '
' tutte le operazioni correlate alla chat , prima vengono chiusi , poi vengono messi in listen '
' in questo fa' eccezione il file transfer che non si mette in listen, vedro'
Private Sub Form_Load()
  login.Icon = icone.Image1.Picture
  Timer_icona.Enabled = True
' all'avvio il programma verifica se si e'connessi ad internet'
  Timer_stato_connessione.Enabled = True
End Sub

Private Sub form_dblclick()
  label15_dblclick
End Sub

Private Sub form_mousedown(Form As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage login.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub

Private Sub label15_dblclick()
 If login.Width < 4860 Then
    Shape1.Visible = True
    login.Height = 9525
    login.Width = 4860
 ElseIf login.Width = 4860 Then
    Timer_ridimensionamento_in_grandezza.Enabled = True
 ElseIf login.WindowState = 2 Then
    Shape1.Visible = True
    login.WindowState = 0
 End If
End Sub

Private Sub Label16_MouseMove(label As Integer, Shift As Integer, X As Single, Y As Single)
Checkremember.BackColor = &H8000000D
End Sub

Private Sub Label17_MouseMove(label As Integer, Shift As Integer, X As Single, Y As Single)
 Checkautoconnect.BackColor = &H8000000D
End Sub

Private Sub Label18_MouseMove(label As Integer, Shift As Integer, X As Single, Y As Single)
 Checkattivasfondi_txtchat.BackColor = &H8000000D
End Sub

Private Sub label19_click()
 Label8_Click
End Sub

Private Sub Label19_MouseMove(label As Integer, Shift As Integer, X As Single, Y As Single)
 shape5.Visible = True
End Sub

Private Sub Label20_Click()
  Shape1.Visible = True
  login.WindowState = 0
 End Sub

Private Sub Label20_MouseMove(label As Integer, Shift As Integer, X As Single, Y As Single)
 Label20.ForeColor = &HFF00FF
End Sub

Private Sub Label21_Click()
 If login.WindowState = 0 Then
    login.WindowState = 2
    Timer_ridimensionamento_in_grandezza.Enabled = True
 ElseIf login.WindowState = 2 Then
    login.WindowState = 0
    Timer_ridimensionamento_in_grandezza.Enabled = True
 End If
 'If Not login.Width = 15360 Then
 '   Timer_ridimensionamento_in_grandezza.Enabled = True
 'End If
End Sub

Private Sub Label8_Click()
 SetParent stato.hwnd, login.Picture12.hwnd
 Picture12.Height = 1260
 Picture12.Width = 1815
 Picture12.Left = 1680
 Picture12.Top = 4850
 stato.Show
 stato.Move 0, 0
 Picture12.Visible = True
End Sub

Private Sub Label8_MouseMove(label As Integer, Shift As Integer, X As Single, Y As Single)
 shape5.Visible = True
End Sub

Private Sub picture12_mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage login.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub

'
'Private Sub Cycle(ByRef banner As Image)
'On Error GoTo ErrHandler
'    bannernumero = bannernumero + 1
'    Txtbannernumero.text = bannernumero
'    Image_banner.Picture = LoadPicture(App.Path & "\banner" & "\immagine" & bannernumero & ".jpg")
'Exit Sub
'ErrHandler:
'    bannernumero = 1
'    Txtbannernumero.text = bannernumero
'    Image_banner.Picture = LoadPicture(App.Path & "\banner" & "\immagine" & bannernumero & ".jpg")
'    Resume Next
'End Sub

Private Sub CandyButton2_Click()
 CandyButton2.Style = XP_ToolBarButton
 Picture11.Top = 10640
 Picture11.Left = 240
End Sub

Private Sub Cmd_fuorissimo_Click()
 browser.browse.Navigate "http://www.fuorissimo.com/"
 SetParent browser.hwnd, login.Picture11.hwnd
 browser.Timer1.Enabled = True
 Picture11.Top = 600
 Picture11.Left = 240
 browser.Show
 browser.Move 0, 0
End Sub

Private Sub Cmd_gnu_Click()
 browser.browse.Navigate "http://www.gnu.org/home.it.html"
 SetParent browser.hwnd, login.Picture11.hwnd
 browser.Timer1.Enabled = True
 Picture11.Top = 600
 Picture11.Left = 240
 browser.Show
 browser.Move 0, 0
End Sub

Private Sub Cmd_planetsourcecode_Click()
 browser.browse.Navigate "http://www.planet-source-code.com/"
 SetParent browser.hwnd, login.Picture11.hwnd
 browser.Timer1.Enabled = True
 Picture11.Top = 600
 Picture11.Left = 240
 browser.Show
 browser.Move 0, 0
End Sub

Private Sub Cmd_scorrere_in_su_Click()
  CandyButton2.Style = Iceblock
  Frame_siti_amici.Top = IIf(Frame_siti_amici.Top <= -Frame_siti_amici.Height, Height, Frame_siti_amici.Top - 800)
End Sub

Private Sub Cmd_sourceforge_Click()
 browser.browse.Navigate "http://sourceforge.net/"
 SetParent browser.hwnd, login.Picture11.hwnd
 browser.Timer1.Enabled = True
 Picture11.Top = 600
 Picture11.Left = 240
 browser.Show
 browser.Move 0, 0
End Sub

Private Sub Cmd_vbfrance_Click()
 browser.browse.Navigate "http://www.vbfrance.com/"
 SetParent browser.hwnd, login.Picture11.hwnd
 browser.Timer1.Enabled = True
 Picture11.Top = 600
 Picture11.Left = 240
 browser.Show
 browser.Move 0, 0
End Sub

Private Sub Cmdgoogle_Click()
 browser.browse.Navigate "http://www.google.it/"
 SetParent browser.hwnd, login.Picture11.hwnd
 browser.Timer1.Enabled = True
 Picture11.Top = 600
 Picture11.Left = 240
 browser.Show
 browser.Move 0, 0
End Sub

Private Sub Cmdscorrere_in_giu_Click()
 Frame_siti_amici.Top = IIf(Frame_siti_amici.Top <= -Frame_siti_amici.Height, Height, Frame_siti_amici.Top + 800)
End Sub

Public Function PLAY_SOUND(Filename As String)
sndPlaySound App.Path & "\" & Filename, SND_ASYNC Or SND_NODEFAULT
End Function

Private Sub Cmdrouter_Click()
 SetParent informazioni_porte.hwnd, login.Picture8.hwnd
 Picture8.Top = 2520
 Picture8.Left = 240
 informazioni_porte.Show
 informazioni_porte.Move 0, 0
End Sub

Private Sub Cmdchiudiprogramma_Click()
 Cmdexit_Click
End Sub

Private Sub Cmdtest_Click()
 SetParent verifica_apertura_porte.hwnd, login.Picture2.hwnd
 Picture2.Top = 600
 Picture2.Left = 240
 verifica_apertura_porte.Show
 verifica_apertura_porte.Move 0, 0
End Sub

Private Sub SaveControlValues()
  Call RegSave(Checkremember, Checkremember.Value)
  Call RegSave(Txtnick, Txtnick.Text)
  Call RegSave(txtPass, txtPass.Text)
  Call RegSave(Txtfrase, Txtfrase.Text)
End Sub

Private Sub Cchiudi_Click()
Frame3.Top = 9480
End Sub

Private Sub Cmdannulla_Click()
Timer_attesa.Enabled = False
  win.Close
  WsMSricevi.Close
  WsPMricevi.Close
  Wsricevifile.Close
  WsPMserverricevi.Close
  Wsricevicomandichat.Close
Cmdannulla.Visible = False
End Sub

Private Sub Cmdbiglietto_Click()
 biglietto_da_visita.Show
End Sub

Private Sub Cmdexit_Click()
 On Error Resume Next
 Picture9.Top = 600
 Picture10.Top = 11280
 Timer_progressbar_uscita.Enabled = True
 Cmddisconnetti_Click
 informazioni_chiusura.Show
 informazioni_chiusura.List1.Visible = True
End Sub

Private Sub Cmdminimizza_Click()
 login.WindowState = 1
End Sub

Private Sub Cmdminimizza_MouseMove(label As Integer, Shift As Integer, X As Single, Y As Single)
 Cmdminimizza.ForeColor = &HFF00FF
End Sub

Private Sub Cmdprofilo_Click()
 Frame3.Top = 5400
 Frame3.Left = 240
End Sub

Private Sub Cmdprofilo_MouseMove(label As Integer, Shift As Integer, X As Single, Y As Single)
 Cmdprofilo.Font.Underline = True
 Cmdprofilo.MouseIcon = pointer_icon.Picture
End Sub

''Private Sub picture_mousemove(label As Integer, Shift As Integer, X As Single, Y As Single)
''  Cmdprofilo.Font.Underline = False
'  Cmdprofilo.MousePointer = vbDefault
'  cmdCreate.Font.Underline = False
'  cmdCreate.MousePointer = vbDefault
'  Label_vediaccount.Font.Underline = False
'  Label_vediaccount.MousePointer = vbDefault
'  Checkremember.BackColor = &H8000000F
'  Checkautoconnect.BackColor = &H8000000F
''  Checkattivasfondi_txtchat.BackColor = &H8000000F
'  cmdLogin.Style = XP_Button
' End Sub

Private Sub Cmdrecuper_Click()
 verifica_recupero_psw_sicurezza.Show
End Sub

Private Sub Cmdritorna_Click()
 Timer_ritorna_in_chat1.Enabled = True
 ' segue il timerritorna in chat2'
End Sub

Private Sub Cmdverifica_Click()
If tentativonumero < 4 Then tentativonumero = tentativonumero + 1
Text4.Text = "tentativo" & tentativonumero
If Txtpassword_sicurezza = psw_sicurezza.Txtpassword_sicurezza Then
   Frame4.Top = 11280
   Frame4.Left = 240
   Frame2.Top = 600
   Frame2.Left = 240
   framelogin.Visible = True
Else
   errore.Show
   ' all'apparire del form avviso regoliamone la posizione di apparizione'
   errore.Top = Me.Top + 3500
   errore.Left = Me.Left + 250
   errore.Labelerrore.Caption = " la password e' errata riprovare"
   ' se si arriva gia' al quarto tentativo errato parte l'allerta'
   If Text4.Text = "tentativo4" Then
      Frame6.Top = 600
      Frame6.Left = 250
      Timer_psw_sicurezza_sbagliata.Enabled = True
      Text4.Text = "" ' il contatore dei tentativi ritorna pulito per dare un anuova possibilita' '
    End If
   Txtpassword_sicurezza.Text = ""
End If

End Sub

Private Sub Form_Resize()
 ' il richiamo del modulo di classe verra' fatto piu' avanti '
  ' per ora lo mettiamo da parte'
  
  'Call frmResize.FormResized(Me)'
    
End Sub

Private Sub cmdCreate_Click()
  If opzioni_segrete.Check5 = 1 Then
     avviso.Show
     avviso.Labelmessaggio.Caption = " hai gia' 2 account registrati, non e' possibile registrarne unaltro"
  Else
     'SetParent regole_canale.hwnd, login.Picture6.hwnd
    ' Picture6.Top = 600
    ' regole_canale.Move 0, 0
     regole_canale.Show
  End If
End Sub

Private Sub cmdCreate_MouseMove(label As Integer, Shift As Integer, X As Single, Y As Single)
 cmdCreate.Font.Underline = True
 cmdCreate.MousePointer = vbCross
End Sub

Private Sub CandyButton_avatar_Click()
  ' SetParent avatar.hwnd, login.Picture5.hwnd
  ' avatar.Show
  ' Picture5.Top = 720
  ' Picture5.Left = 1320
  ' avatar.Move 0, 0
  avatar.Show 1
 End Sub

Private Sub cmdLogin_Click()
 Static connect As Integer
 If cmdLogin.Caption = "accedi" Then
    If connect = 0 Then
        connect = 1
        lblStatus = "Connecting..."
        thetime = GetTickCount
        Do While (win.State <> 7) And thetime > GetTickCount - 10000
            win.Close
            win.connect Txtip.Text, 12584
            DoEvents
        Loop
        If win.State = 7 Then
            If Txtnick <> "" And txtPass <> "" Then
                lblStatus = "Sending username/password"
                win.SendData "User-" & Txtnick & "\-"
                
                lblStatus = "Verifying..."
                win.SendData "Pass-" & txtPass & "\-"
            End If
        Else
            win.Close
            MsgBox "Couldn't connect to server!", , "Error!"
            lblStatus = "Couldn't connect!"
        End If
        connect = 0
    End If
    
        cmdLogin.Caption = "annulla"
 ElseIf cmdLogin.Caption = "annulla" Then
        Cmdannulla_Click
        Picture12.Visible = False
        Picture12.Height = 9495
        Picture12.Width = 4815
        Picture12.Top = 0
        Picture12.Left = 5640
        Checkremember.Visible = True
        Label16.Visible = True
        Checkautoconnect.Visible = True
        Label17.Visible = True
        Checkattivasfondi_txtchat.Visible = True
        Label18.Visible = True
        animazione_connessione.Visible = False
        cmdLogin.Caption = "accedi"
 End If
 
End Sub

Private Sub cmdLogin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 cmdLogin.Style = Iceblock
End Sub

Private Sub Cmddisconnetti_Click()
 Ws.Close
 win.Close
 WsMSricevi.Close
 WsPMricevi.Close
 TrayDelete
 Timer_unload1.Enabled = True
End Sub

Private Sub Option1_Click()

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

Private Sub Label_tempo_di_connessione_mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage login.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub

Private Sub Label_vediaccount_Click()
 If psw_account.Txtpsw_account.Text = "" Then
    avviso.Show
    avviso.Labelmessaggio.Caption = "devi creare la password per proteggere i tuoi account"
    psw_account.Show
 Else
    account.Show
 End If
End Sub

Private Sub Label_vediaccount_MouseMove(label As Integer, Shift As Integer, X As Single, Y As Single)
 Label_vediaccount.Font.Underline = True
 Label_vediaccount.MousePointer = vbCross
End Sub

Private Sub Picture1_Click()
Picture = ImageList1.ListImages.Item(1).Picture
End Sub

Private Sub Timer_attesa_Timer()
If Not Ws.State = sckConnected Then
   Ws.connect Txtip.Text, 1000 ' Connects to the server
End If
 If Checkremember = 1 Then
  Call RegSave(Checkremember, Checkremember.Value) ' salviamo la scelta fatta chiamando il regsave'
                                                   ' che si trovanel modulo salva impostazioni'
  Call RegSave(Txtnick, Txtnick.Text)
  Call RegSave(txtPass, txtPass.Text)
  Call RegSave(Txtfrase, Txtfrase.Text)
End If
 Call RegSave(Checkautoconnect, Checkautoconnect.Value)
 Timer_attesa.Enabled = False
End Sub

Private Sub Timer_bannaggio_Timer()
login.Visible = False
Timer_bannaggio.Enabled = False
End Sub

Private Sub Timer_caricamento_componenti_allavvio_Timer()
Label1 = Ws.LocalIP
Anim2.AnimatedGifPath = App.Path & "\animazioni" & "\MS" & "\immagine4.gif"
  frmResize.KeepRatio = True
  frmResize.FontResize = True
  Call frmResize.InitializeResizer(Me)
Cmddisconnetti.Enabled = False
 Me.Show
    sent = False
    addingit = False
    Txtnick = GetSetting("IMer", "Saved info", "Username")
    If Txtnick <> "" Then
        txtPass.SetFocus
    End If
 Checkremember.Value = RegLoad(Checkremember) ' all'avvio ricordiamo se il check era 1 o  0 '
 Checkautoconnect.Value = RegLoad(Checkautoconnect)
 Txtnick = RegLoad(Txtnick)
 txtPass = RegLoad(txtPass)
 Txtfrase = RegLoad(Txtfrase)
 psw_sicurezza.Txtpassword_sicurezza = RegLoad(psw_sicurezza.Txtpassword_sicurezza)
 opzioni.Check5.Value = RegLoad(opzioni.Check5)
 psw_account.Txtpsw_account = RegLoad(psw_account.Txtpsw_account)
 psw_account.Text1 = RegLoad(psw_account.Text1)
 sistema_di_crediti.Timer_salvataggio.Enabled = True ' all'avvio abilitiamo il timer per il salvataggio automatico'
 sistema_di_crediti.Timer_caricamento.Enabled = True 'all'avvio vengono caricate le impostazioni salvate'
 caricamento.Text2.Visible = True
 Timer_caricamento_componenti_allavvio.Enabled = False
 Timer_check.Enabled = True
End Sub

Private Sub Timer_check_Timer()
If opzioni_segrete.Check9 = 1 Then
   Anim1.Visible = True
   Anim1.AnimatedGifPath = App.Path & "\immagini varie" & "\moderatore.jpg"
End If
If Checkautoconnect = 1 Then
 cmdLogin_Click
End If
If opzioni.Check1 = 1 Then
    avviso_connessione.Visible = False
  End If
If opzioni.Check4 = 1 Then
   framelogin.Visible = False
   Frame2.Top = 9480
   Frame4.Top = 720
ElseIf opzioni.Check5 = 1 Then
   popup.Visible = False
End If
biglietto_da_visita.biglietto = LoadText(App.Path & "\informazioni utente\biglietto da visita\biglietto.txt")
caricamento.Text3.Visible = True
Anim4.AnimatedGifPath = App.Path & "\immagini varie" & "\immagine3.gif"
Timer_grafica_in_entrata.Enabled = True
Timer_tempo_di_connessione.Enabled = True
StartTime = Now
Timer_check.Enabled = False
End Sub

Private Sub Timer_grafica_in_entrata_Timer()
 trasparenza_per_login.Height = login.Height
 trasparenza_per_login.Width = login.Width
 trasparenza_per_login.Top = login.Top
 trasparenza_per_login.Left = login.Left
 trasparenza_per_login.Show
 trasparenza_per_login.Timer_progressbar_per_trasparenza.Enabled = True
 Picture12.Top = 0
 Picture12.Left = 7080
 Unload caricamento
 Picture12.Visible = False
 Timer_grafica_in_entrata.Enabled = False
End Sub

Private Sub Timer_icona_Timer()
 TrayAdd hwnd, Me.Icon, "System Tray", MouseMove
 Timer_icona.Enabled = False
End Sub

Private Sub Timer_lmage1_Timer()
Static frame As Integer

frame = frame + 1

If frame > anim.ListImages.Count Then frame = 1

Image1.Picture = anim.ListImages(frame).Picture

End Sub


Private Sub Timer_login_chat1_Timer()
  login.Ws.SendData "@CONNECT:" & Txtnick.Text & "   " & Label1 & "   " & avatar.Txtavatar.Text
  Cmdannulla.Visible = False
  Timer_login_chat1.Enabled = False
  Timer_login_chat2.Enabled = True
End Sub

Private Sub Timer_login_chat2_Timer()
 Frame2.Top = 5400
 Frame2.Left = 240
 Picture12.Visible = False
 Picture12.Height = 9495
 Picture12.Width = 4815
 Picture12.Top = 0
 Picture12.Left = 5640
 Checkremember.Visible = True
 Label16.Visible = True
 Checkautoconnect.Visible = True
 Label17.Visible = True
 Checkattivasfondi_txtchat.Visible = True
 Label18.Visible = True
 animazione_connessione.Visible = False
 'frame8.Top = 600'
 login.Visible = False
 Unload avviso_connessione
 PLAY_SOUND "connected"
 chat_style.Show
 chat.Visible = True
 cmdLogin.Caption = "accedi"
 Timer_login_chat2.Enabled = False
 Timer_login_im.Enabled = True
End Sub

Private Sub Timer_progressbar_Timer()
  ProgressBar1.Value = ProgressBar1.Value + 1
  Label7.Caption = Label7.Caption + 1
  If ProgressBar1.Value = 100 Then
  Timer_progressbar.Enabled = False
  Label7.Caption = "100"
  End If
End Sub

Private Sub Timer_login_im_Timer()
 login.WindowState = 0
 SetParent frmBuddyList.hwnd, login.Picture12.hwnd
 frmBuddyList.Show
 frmBuddyList.Move 0, 0
 Picture12.Top = 0
 Picture12.Left = 0
 Picture12.Height = login.Height
 Picture12.Width = login.Width
 Picture12.Visible = True
 Timer_login_im.Enabled = False
End Sub

Private Sub Timer_psw_sicurezza_sbagliata_Timer()
 Frame6.Left = 7080
  Timer_psw_sicurezza_sbagliata.Enabled = False
End Sub

Private Sub Timer_ridimensionamento_in_grandezza_Timer()
    Shape1.Visible = False
    'login.WindowState = 2'
    Image1.Height = 615
    Image1.Width = 645
    Image1.Left = 0
    Image1.Top = 0
    Label15.Height = 255
    Label15.Width = 975
    'Label15.Left = 6600
    Label15.Left = 840
    Label_tempo_di_connessione.Height = 255
    Label_tempo_di_connessione.Width = 1935
    'Label_tempo_di_connessione.Left = 7440
    Label_tempo_di_connessione.Left = 1680
    Cmdminimizza.Height = 255
    Cmdminimizza.Width = 255
    'Cmdminimizza.Left = 9120
   
    Shape6.Height = 210
    Shape6.Width = 240
    'Shape6.Left = 9120
    Label20.Height = 255
    Label20.Width = 255
    'Label20.Left = 9480
    Shape7.Height = 210
    Shape7.Width = 240
    'Shape7.Left = 9840
    Label21.Height = 255
    Label21.Width = 255
    'Label21.Left = 9840
    Cmdexit.Height = 255
    Cmdexit.Width = 255
    'Cmdexit.Left = 10200
    CandyButton_avatar.Height = 375
    CandyButton_avatar.Width = 1095
    'CandyButton_avatar.Left = 9000
    Picavatar.Height = 1455
    Picavatar.Width = 1455
    'Picavatar.Left = 7440
    Labelnick.Height = 255
    Labelnick.Width = 1215
    'Labelnick.Left = 6840
    Txtnick.Height = 285
    Txtnick.Width = 2655
    'Txtnick.Left = 6840
    Label2.Height = 255
    Label2.Width = 855
    'Label2.Left = 6840
    txtPass.Height = 285
    txtPass.Width = 2655
    'txtPass.Left = 6840
    Labelfrase.Height = 255
    Labelfrase.Width = 855
    'Labelfrase.Left = 6840
    Txtfrase.Height = 375
    Txtfrase.Width = 2655
    'Txtfrase.Left = 6840
    Label7.Height = 255
    Label7.Width = 615
    'Label7.Left = 6840
    shape5.Height = 255
    shape5.Width = 1455
    'shape5.Left = 7440
    Label8.Height = 255
    Label8.Width = 975
    'Label8.Left = 7560
    Label19.Height = 255
    Label19.Width = 255
    'Label19.Left = 8640
    Checkremember.Height = 255
    Checkremember.Width = 255
    'Checkremember.Left = 6840
    Label16.Height = 255
    Label16.Width = 3255
    'Label16.Left = 7200
    Checkautoconnect.Height = 255
    Checkautoconnect.Width = 255
    'Checkautoconnect.Left = 6840
    Label17.Height = 255
    Label17.Width = 3135
    'Label17.Left = 7200
    Checkattivasfondi_txtchat.Height = 255
    Checkattivasfondi_txtchat.Width = 255
    'Checkattivasfondi_txtchat.Left = 6840
    Label18.Height = 255
    Label18.Width = 2055
    'Label18.Left = 7200
    cmdLogin.Height = 375
    cmdLogin.Width = 855
    'cmdLogin.Left = 7080
    Cmdprofilo.Height = 255
    Cmdprofilo.Width = 1095
    'Cmdprofilo.Left = 5880
    cmdCreate.Height = 255
    cmdCreate.Width = 1575
    'cmdCreate.Left = 5880
    Label_vediaccount.Height = 255
    Label_vediaccount.Width = 1695
    'Label_vediaccount.Left = 5880
    Cmdrouter.Height = 255
    Cmdrouter.Width = 855
    'Cmdrouter.Left = 9240
    Cmdtest.Height = 255
    Cmdtest.Width = 1455
    'Cmdtest.Left = 9240
    Anim1.Height = 1215
    Anim1.Width = 1455
    'Anim1.Left = 7320
    Anim4.Height = 735
    Anim4.Width = 855
    'Anim4.Left = 9240
   Timer_ridimensionamento_in_grandezza.Enabled = False
End Sub

Private Sub Timer_ridimensionamento_Timer()
    Image1.Height = 615
    Image1.Width = 645
    
    Label15.Height = 255
    Label15.Width = 975
    
    Label_tempo_di_connessione.Height = 255
    Label_tempo_di_connessione.Width = 1935
   
    Cmdminimizza.Height = 255
    Cmdminimizza.Width = 255
    
    Shape6.Height = 210
    Shape6.Width = 240
    
    Label20.Height = 255
    Label20.Width = 255
    
    Shape7.Height = 210
    Shape7.Width = 240
    
    Label21.Height = 255
    Label21.Width = 255
    
    Cmdexit.Height = 255
    Cmdexit.Width = 255
    
    CandyButton_avatar.Height = 375
    CandyButton_avatar.Width = 1095
    
    Picavatar.Height = 1455
    Picavatar.Width = 1455
    
    Labelnick.Height = 255
    Labelnick.Width = 1215
    
    Txtnick.Height = 285
    Txtnick.Width = 2655
    
    Label2.Height = 255
    Label2.Width = 855
    
    txtPass.Height = 285
    txtPass.Width = 2655
    
    Labelfrase.Height = 255
    Labelfrase.Width = 855
    
    Txtfrase.Height = 375
    Txtfrase.Width = 2655
    
    Label7.Height = 255
    Label7.Width = 615
    
    shape5.Height = 255
    shape5.Width = 1455
    
    Label8.Height = 255
    Label8.Width = 975
    
    Label19.Height = 255
    Label19.Width = 255
    
    Checkremember.Height = 255
    Checkremember.Width = 255
    
    Label16.Height = 255
    Label16.Width = 3255
    
    Checkautoconnect.Height = 255
    Checkautoconnect.Width = 255
    
    Label17.Height = 255
    Label17.Width = 3135
   
    Checkattivasfondi_txtchat.Height = 255
    Checkattivasfondi_txtchat.Width = 255
    
    Label18.Height = 255
    Label18.Width = 2055
    
    cmdLogin.Height = 375
    cmdLogin.Width = 855
    
    Cmdprofilo.Height = 255
    Cmdprofilo.Width = 1095
    
    cmdCreate.Height = 255
    cmdCreate.Width = 1575
   
    Label_vediaccount.Height = 255
    Label_vediaccount.Width = 1695
    
    Cmdrouter.Height = 255
    Cmdrouter.Width = 855
    
    Cmdtest.Height = 255
    Cmdtest.Width = 1455
    
    Anim1.Height = 1215
    Anim1.Width = 1455
    
    Anim4.Height = 735
    Anim4.Width = 855
    Timer_ridimensionamento.Enabled = False
End Sub

Private Sub Timer_ritardo_Timer()
Wsricevicomandichat.SendData biglietto_da_visita.biglietto.Text
Timer_ritardo.Enabled = False
End Sub

Private Sub Timer_richiamo_salvataggio_opzioni_segrete_Timer()
 opzioni_segrete.Check1.Value = RegLoad(opzioni_segrete.Check1)
 opzioni_segrete.Check2.Value = RegLoad(opzioni_segrete.Check2)
 opzioni_segrete.Check3.Value = RegLoad(opzioni_segrete.Check3)
 opzioni_segrete.Check4.Value = RegLoad(opzioni_segrete.Check4)
 opzioni_segrete.Check5.Value = RegLoad(opzioni_segrete.Check5)
 opzioni_segrete.Check6.Value = RegLoad(opzioni_segrete.Check6)
 opzioni_segrete.Check7.Value = RegLoad(opzioni_segrete.Check7)
 opzioni_segrete.Check8.Value = RegLoad(opzioni_segrete.Check8)
 opzioni_segrete.Check9.Value = RegLoad(opzioni_segrete.Check9)
caricamento.Text1.Visible = True
Timer_richiamo_salvataggio_opzioni_segrete.Enabled = False
 Timer_caricamento_componenti_allavvio.Enabled = True
End Sub

Private Sub Timer_ritorna_in_chat1_Timer()
 login.Ws.SendData "@CONNECT:" & Txtnick.Text & "   " & Label1 & "   " & avatar.Txtavatar.Text
 SetParent avviso_ritorno_in_chat.hwnd, login.Picture4.hwnd
 Picture4.Top = 3000
 Picture4.Left = 1000
 avviso_ritorno_in_chat.Show
 avviso_ritorno_in_chat.Move 0, 0
 Timer_ritorna_in_chat1.Enabled = False
 Timer_ritorna_in_chat2.Enabled = True
End Sub

Private Sub Timer_ritorna_in_chat2_Timer()
 Unload avviso_ritorno_in_chat
 'Frame8.Top = 600'
 Picture4.Top = 9600
 login.Visible = False
 chat_style.Show
 chat.Show
 Timer_ritorna_in_chat2.Enabled = False
End Sub

' all'avvio nel form_load , facciamo partire il timer perverificare lo stato della'
' connessione , se non si e' connessi ad internet non parte il programma'
Private Sub Timer_stato_connessione_Timer()
 SetParent stato_connessione.hwnd, login.Picture12.hwnd
 Picture12.Visible = True
 Picture12.Top = 0
 Picture12.Left = 0
 stato_connessione.Show
 stato_connessione.Move 0, 0
 timer_stato_connessione2.Enabled = True
 Timer_stato_connessione.Enabled = False
End Sub

Private Sub timer_stato_connessione2_Timer()
 If stato_connessione.txtStatus.Text = "siete connessi ad internet" Then
  caricamento.Anim3.AnimatedGifPath = App.Path & "\immagini varie" & "\immagine2" & ".gif"
  stato_connessione.Visible = False
  ' all'avvio del programma la picture12 assumera' questa posizione'
  SetParent caricamento.hwnd, login.Picture12.hwnd
  caricamento.Show
  caricamento.Move 0, 0
  Timer_winsock_close.Enabled = True
 ElseIf stato_connessione.txtStatus.Text = "non siete connessi ad internet" Then
  Picture12.Top = 0
  Picture12.Left = 0
 End If
 timer_stato_connessione2.Enabled = False
End Sub

Private Sub Timer_tempo_di_connessione_Timer()
 Label_tempo_di_connessione.Caption = Format$(Now - StartTime, "hh:mm:ss")
End Sub

Private Sub Timer_unload_frmbuddylist_Timer()
 frmAdd.Hide
 frmBuddyList.Hide
 frmBuddyList.lstBuddy.Clear
 frmBuddyList.lstOffline.Clear
 timesgotinfo = 0
 frmCreate.Hide
 frmIM.Hide
 frmInfo.Hide
 frmSetInfo.Hide
 Unload frmBuddyList
 Picture12.Visible = False
 Picture12.Height = 9495
 Picture12.Width = 4815
 Picture12.Top = 0
 Picture12.Left = 5640
 Timer_unload_frmbuddylist.Enabled = False
End Sub

'con questo timer chiudiamo tutti i winsock'
Private Sub Timer_unload1_Timer()
 Unload account
 Unload agenda
 Unload animazioni_chat
 Unload animazioni_flash_chat
 Unload animazioni_MSinvio
 Unload animazioni_MSricevi
 Unload avatar
 Unload avatar_animati
 Unload avviso
 Unload avviso_chiusura
 Unload avviso_connessione
 Unload avviso_ritorno_in_chat
 Unload biglietto_da_visita
 Unload bigsmile
 Unload blocca_sblocca
 Unload block_privat
 Unload Calcolatrice
 Unload cambianick
 Unload cercautente
 Timer_unload1.Enabled = False
 Label9.Visible = True
 informazioni_chiusura.List2.Visible = True
 Timer_unload2.Enabled = True ' abilitiamo il secondo timer per la chiusura'
End Sub

Private Sub Timer_unload2_Timer()
 Unload crediti
 Unload criptafile
 Unload decriptafile
 Unload errore
 Unload esegui_animazioni_chat_inviate
 Unload esegui_animazioni_chat_ricevute
 Unload esegui_animazioni_flash_chat_inviate
 Unload esegui_animazioni_flash_chat_ricevute
 Unload esegui_animazioni_MSinvio
 Unload esegui_animazioni_MSricevi
 Unload esegui_biglietto_da_visita
 Unload eventi
 Unload FILEinvia
 Unload FILEricevi
 Timer_unload2.Enabled = False
 Label10.Visible = True
 Timer_unload3.Enabled = True
 informazioni_chiusura.List3.Visible = True
End Sub

Private Sub Timer_unload3_Timer()
 Unload frmAdd
 Unload frmBuddyList
 Unload frmCreate
 Unload frmIM
 Unload frmInfo
 Unload frmRyCamV2
 Unload frmRyCamv2a
 Unload frmRyCamV2b
 Unload frmSetInfo
 Unload informazioni_porte
 Unload informazioniutente
 Unload invia_comandi_chat
 Unload licenza
 Unload link
 Unload messaggio_globale
 Unload messaggiomassa
 Unload MOD_chat
 Unload MOD_login
 Unload MODERAZIONE
 Timer_unload3.Enabled = False
 Label11.Visible = True
 Timer_unload4.Enabled = True
 informazioni_chiusura.List4.Visible = True
End Sub

Private Sub Timer_unload4_Timer()
 Unload nudge
 Unload opzioni
 Unload opzioni_segrete
 Unload PMinvio
 Unload PMricevi
 Unload PMserverinvio
 Unload PMserverricevi
 Unload popup
 Unload PRIVILEGI
 Unload profiloutente
 Unload psw_account
 Unload psw_sicurezza
 Unload recupero_password_sicurezza
 Timer_unload4.Enabled = False
 Label12.Visible = True
 Timer_unload5.Enabled = True
 informazioni_chiusura.List5.Visible = True
End Sub

Private Sub Timer_unload5_Timer()
 Unload regole_canale
 Unload riabilita_privat
 Unload ricevi_comandi_chat
 Unload risposte
 Unload risposte_veloci
 Unload sfondi
 Unload sfondi_chatprivata
 Unload sfondi_txtchat
 Unload sistema_di_crediti
 Unload smyle
 Unload tempo
 Timer_unload5.Enabled = False
 Label13.Visible = True
 informazioni_chiusura.List6.Visible = True
 Timer_unload6.Enabled = True
End Sub

Private Sub Timer_unload6_Timer()
 Label14.Visible = True
 Unload verifica_apertura_porte
 Unload verifica_recupero_psw_sicurezza
 Unload webbrowser_banner
 Timer_unload6.Enabled = False
 Unload informazioni_chiusura
 Timer_unload7.Enabled = True
End Sub

Private Sub Timer_unload7_Timer()
 'non occorre dareunload al form login perche' dando end si azzera tutto'
 Unload MSinvio
 Unload MSinvio_style
 Unload MSricevi
 Unload MSricevi_styleE
 Unload login
 End
 Timer_unload7.Enabled = False
End Sub


'allavvio chiudiamo tutti gli winsock per avere un avvio pulito'
' se percaso in un precedente utilizzo si e' avuta una chiusura non adeguata'
' potrebbe aversi un errore di indirizzo in uso, cosi' partiamo a winsock chiusi'
Private Sub Timer_winsock_close_Timer()
  WsMSricevi.Close
  WsPMricevi.Close
  Wsricevifile.Close
  WsPMserverricevi.Close
  Wsricevicomandichat.Close
  Label14.Visible = True
 Timer_winsock_close.Enabled = False
 Timer_richiamo_salvataggio_opzioni_segrete.Enabled = True
End Sub

Private Sub Timer_winsock_listen_Timer()
  If WsMSricevi.State <> sckClosed Then
     WsMSricevi.Listen
  End If
  If WsPMricevi.State <> sckClosed Then
     WsPMricevi.Listen
  End If
  If WsPMserverricevi.State <> sckClosed Then
     WsPMserverricevi.Listen
  End If
  If Wsricevicomandichat.State <> sckClosed Then
     Wsricevicomandichat.Listen
  End If
 Timer_winsock_listen.Enabled = False
End Sub

' informazioni dopo la connessione e prima di accedere'
' alla chat'
Private Sub Ws_Connect()
 avviso_connessione.Label4.Visible = True
 Timer_login_chat1.Enabled = True
 Timer_winsock_listen.Enabled = True
End Sub

'                                                  '
'   coded by h2o113 of p2pforum                    '
'                                                  '
'   la lista degli utenti verra' spedita al server '

Private Sub Ws_DataArrival(ByVal bytesTotal As Long)
Dim Nick, Frase, Messaggio As String
Dim txt As String
Dim Splittati() As String
Ws.GetData txt, vbString
If Mid(txt, 1, 6) = "@LISTC" Then
    chat.listusers.Clear
    Splittati() = Split(Mid(txt, 7), "@")
    For i = LBound(Splittati) + 1 To UBound(Splittati)
        Me.Caption = Splittati(i)
            chat.listusers.AddItem Mid(Splittati(i), InStr(1, Splittati(i), "LIST:") + 5)
   Next i
    Exit Sub
End If
If InStr(1, txt, Chr(127) & "cambionick:") > 0 Then
Nick = Mid(txt, InStr(1, txt, Chr(127) & "nick:") + 6, InStr(1, txt, Chr(127) & "cambionick:") - (InStr(1, txt, Chr(127) & "nick:") + 6))
Frase = Mid(txt, InStr(1, txt, Chr(127) & "cambionick:") + 12)
Messaggio = ""
ChatMessage Nick, Frase, Messaggio
Exit Sub
End If


Nick = Mid(txt, InStr(1, txt, Chr(127) & "nick:") + 6, InStr(1, txt, Chr(127) & "frase:") - (InStr(1, txt, Chr(127) & "nick:") + 6))
Frase = Mid(txt, InStr(1, txt, Chr(127) & "frase:") + 7, InStr(1, txt, Chr(127) & vbCrLf) - (InStr(1, txt, Chr(127) & "frase:") + 7))
Messaggio = Mid(txt, InStr(1, txt, Chr(127) & vbCrLf) + 3)
'chat.txtchat.Text = chat.txtchat.Text + txt + vbCrLf ' il txtchat contiene sia i messaggi inviati che ricevuti'

ChatMessage Nick, Frase, Messaggio
chat.txtSend.Text = ""
End Sub

Private Sub WsPMricevi_ConnectionRequest(ByVal requestID As Long)
    If WsPMricevi.State <> sckClosed Then WsPMricevi.Close
    WsPMricevi.Accept requestID ' accetta la connessione '
    PMricevi.Show
    Newalert.Message = "ti e' stato recapitato un messaggio privato " '
    Newalert.Title = "arrivo PM"                            ' un popup ci avvisa che e' arrivato'
    Newalert.MSN6                                             ' un  messaggio privato'
End Sub

 ' se la connessione con il client termina il winsock ritorna il listen '
Private Sub WsPMricevi_close()
If WsPMricevi.State <> sckClosed Then WsPMricevi.Close
 WsPMricevi.Listen ' il winsock ritorna il listen '
End Sub

Private Sub WsPMricevi_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
WsPMricevi.GetData Data
PMricevi.Txtmessaggio.Text = Data
End Sub

Private Sub WsMSricevi_ConnectionRequest(ByVal requestID As Long)
    
    If WsMSricevi.State <> sckClosed Then WsMSricevi.Close
    WsMSricevi.Accept requestID ' accetta la connessione dall'amico'
Newalert.Message = "e' stata richiesta una chat privata " '
Newalert.Title = "inizio chat"                            ' un popup ci avvisa che qualcuno'
Newalert.MSN6                                             ' vuole iniziare una chat privata'
    'e' attiva una connessione, ed il form MSricevi viene fatto apparire'
  MSricevi_styleE.Show
End Sub

 ' sei il client si disconnette ora il problema della riconessione e' risolto'
 ' in quanto il wiinsock viene rimesso in listen '
Private Sub WsMSricevi_close()
If WsMSricevi.State <> sckClosed Then WsMSricevi.Close
 WsMSricevi.Listen
End Sub

Private Sub WsMSricevi_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
WsMSricevi.GetData Data
If Data = "(((($$$$ nudge ££££)))))" Then ' stabiliamo che se data e' uguale a questi insieme di simboli'
nudge.TimernudgeMSricevi.Enabled = True   ' chi li riceve attiva il nudge'
ElseIf Data = "animazione1" Then
 esegui_animazioni_MSricevi.Show
 esegui_animazioni_MSricevi.Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\MS" & "\immagine1.gif"
ElseIf Data = "animazione2" Then
 esegui_animazioni_MSricevi.Show
 esegui_animazioni_MSricevi.Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\MS" & "\immagine2.gif"
ElseIf Data = "animazione3" Then
 esegui_animazioni_MSricevi.Show
 esegui_animazioni_MSricevi.Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\MS" & "\immagine3.gif"
ElseIf Data = "animazione4" Then
 esegui_animazioni_MSricevi.Show
 esegui_animazioni_MSricevi.Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\MS" & "\immagine4.gif"
ElseIf Data = "sfondo1" Then
 MSricevi_styleE.Anim1.AnimatedGifPath = App.Path & "\sfondi" & "\privatchat" & "\immagine1.jpg"
 MakeTransparent MSricevi.hwnd, 200
ElseIf Data = "sfondo2" Then
 MSricevi_styleE.Anim1.AnimatedGifPath = App.Path & "\sfondi" & "\privatchat" & "\immagine2.jpg"
 MakeTransparent MSricevi.hwnd, 200
ElseIf Data = "sfondo3" Then
 MSricevi_styleE.Anim1.AnimatedGifPath = App.Path & "\sfondi" & "\privatchat" & "\immagine3.jpg"
 MakeTransparent MSricevi.hwnd, 200
ElseIf Data = "sfondo4" Then
 MSricevi_styleE.Anim1.AnimatedGifPath = App.Path & "\sfondi" & "\privatchat" & "\immagine4.jpg"
 MakeTransparent MSricevi.hwnd, 200
ElseIf Data = "sfondo5" Then
 MSricevi_styleE.Anim1.AnimatedGifPath = App.Path & "\sfondi" & "\privatchat" & "\immagine5.jpg"
 MakeTransparent MSricevi.hwnd, 200
ElseIf Data = "sfondo6" Then
 MSricevi_styleE.Anim1.AnimatedGifPath = App.Path & "\sfondi" & "\privatchat" & "\immagine6.jpg"
 MakeTransparent MSricevi.hwnd, 200
End If
 MSricevi.Listms.AddItem Data
End Sub

Private Sub Wsricevifile_ConnectionRequest(ByVal requestID As Long)
If Wsricevifile.State <> sckClosed Then Wsricevifile.Close 'if state is closed, then close the socket
    Wsricevifile.Accept requestID ' accept connection from client
    MsgBox " e' stata stabilita una connessione per l'invio dei file"
   FILEricevi.Show
 FILEricevi.Listavviso.AddItem " qualcuno ti sta' inviando il file"
End Sub

Private Sub wsricevifile_close()
If Wsricevifile.State <> sckClosed Then Wsricevifile.Close
 Wsricevifile.Listen
FILEricevi.Listavviso.Clear
End Sub

Private Sub Wsricevifile_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
Dim StrSplited() As String
Dim StrFilename As String
Dim LngFileSize As Long

Wsricevifile.GetData strData 'getdata'
StrSplited = Split(strData, "|") 'we split data (delimiter = |)'

If strData = "E" Then ' if it's the end of a transfert'
    Close #2 'we close the dest file'
    BlnTflag = False 'we update the flag'
    Exit Sub
End If

If BlnTflag = False Then ' if we're not in a file transfert'
'transfert initialisation: flag update, file name get, we open a free file, then we ask for the first chunk'
    If StrSplited(0) = "Transfert" Then
        StrFilename = StrSplited(1)
        BlnTflag = True
        Wsricevifile.SendData "S"
        If Dir(App.Path & "\" & StrFilename) <> "" Then ' we erase the file if it exists'
            Kill (App.Path & "\" & StrFilename)
        End If
        
        Open App.Path & "\" & StrFilename For Binary As #2
    End If
Else
    Put #2, LOF(2) + 1, strData ' we write data at the end of the file'
    Wsricevifile.SendData "N" 'we ask for another chunk'
End If
End Sub

Private Sub WsPMserverricevi_ConnectionRequest(ByVal requestID As Long)
    If WsPMserverricevi.State <> sckClosed Then WsPMserverricevi.Close
    WsPMserverricevi.Accept requestID ' accetta la connessione '
    PMserverricevi.Show
    Newalert.Message = "ti e' stato recapitato un messaggio privato da parte del server" '
    Newalert.Title = "arrivo PM"                            ' un popup ci avvisa che e' arrivato'
    Newalert.MSN6                                             ' un  messaggio privato'
End Sub

 ' se la connessione con il client termina il winsock ritorna il listen '
Private Sub WsPMserverricevi_close()
If WsPMserverricevi.State <> sckClosed Then WsPMserverricevi.Close
 WsPMserverricevi.Listen ' il winsock ritorna il listen '
End Sub

Private Sub WsPMserverricevi_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
WsPMserverricevi.GetData Data
PMserverricevi.Txtmessaggio.Text = Data
' COMANDI DAL SERVER'
If Data = "(((end)))" Then
 login.Visible = True
login.Ws.SendData "@DISCONNECT:" & login.Txtnick.Text & "   " & login.Label1 & "   " & avatar.Txtavatar.Text
chat_style.Visible = False
chat.Visible = False
ElseIf Data = "(((ban)))" Then
 login.Visible = True
 Timer_bannaggio.Enabled = True
ElseIf Data = "(((block privat)))" Then
 chat.CmdMS.Visible = False
 chat.CmdPM.Visible = False
 chat.Cmdinviafile.Visible = False
 chat.Cmdwebcam.Visible = False
 chat.cmdIM.Visible = False
ElseIf Data = "(((lock chat)))" Then
 chat.txtChat.Locked = True
 chat.listusers.Enabled = False
ElseIf Data = "(((block-keyboard-mouse)))" Then
 BlockInput True
ElseIf Data = "(((sblock-keyboard-mouse)))" Then
 BlockInput False
' richiamiamo il comando per diventare moderatore'
ElseIf Data = "(((richiesta moderatore)))" Then
 opzioni_segrete.Check9 = 1
 avviso.Show
 avviso.Labelmessaggio.Caption = " al prossimo riavvia sarai moderatore"
ElseIf Data = "(((rimuovi moderatore)))" Then
 opzioni_segrete.Check9 = 0
 avviso.Show
 avviso.Labelmessaggio.Caption = " la tua moderazione e' stata sospesa"
End If
End Sub

Private Sub Wsricevicomandichat_ConnectionRequest(ByVal requestID As Long)
  If Wsricevicomandichat.State <> sckClosed Then Wsricevicomandichat.Close 'if state is closed, then close the socket
      Wsricevicomandichat.Accept requestID
       ricevi_comandi_chat.Show
  End Sub
 
 ' se la connessione con il client termina il winsock ritorna il listen '
Private Sub Wsricevicomandichat_close()
If Wsricevicomandichat.State <> sckClosed Then Wsricevicomandichat.Close
 Wsricevicomandichat.Listen ' il winsock ritorna il listen '
End Sub

Private Sub Wsricevicomandichat_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
Wsricevicomandichat.GetData Data ' store data in Data
 If Data = "(((nudge_chat)))" Then
  nudge.Timernudge_chat.Enabled = True
  PLAY_SOUND "nudge"
 ElseIf Data = "(((richiesta biglietto da visita)))" Then
  Wsricevicomandichat.SendData "(((invio biglietto)))"
  Timer_ritardo.Enabled = True
 ElseIf Data = "animazione1" Then
  esegui_animazioni_chat_ricevute.Show
  esegui_animazioni_chat_ricevute.Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\chat" & "\immagine1" & ".gif"
 ElseIf Data = "animazione2" Then
  esegui_animazioni_chat_ricevute.Show
  esegui_animazioni_chat_ricevute.Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\chat" & "\immagine2" & ".gif"
 ElseIf Data = "animazione3" Then
  esegui_animazioni_chat_ricevute.Show
  esegui_animazioni_chat_ricevute.Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\chat" & "\immagine3" & ".gif"
 ElseIf Data = "animazione4" Then
  esegui_animazioni_chat_ricevute.Show
  esegui_animazioni_chat_ricevute.Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\chat" & "\immagine4" & ".gif"
 ElseIf Data = "animazione5" Then
  esegui_animazioni_chat_ricevute.Show
  esegui_animazioni_chat_ricevute.Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\chat" & "\immagine5" & ".gif"
  esegui_animazioni_chat_ricevute.Show
  esegui_animazioni_chat_ricevute.Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\chat" & "\immagine6" & ".gif"
 ElseIf Data = "animazione7" Then
  esegui_animazioni_chat_ricevute.Show
  esegui_animazioni_chat_ricevute.Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\chat" & "\immagine7" & ".gif"
 ElseIf Data = "animazione8" Then
  esegui_animazioni_chat_ricevute.Show
  esegui_animazioni_chat_ricevute.Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\chat" & "\immagine8" & ".gif"
 ElseIf Data = "animazione9" Then
  esegui_animazioni_chat_ricevute.Show
  esegui_animazioni_chat_ricevute.Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\chat" & "\immagine9" & ".gif"
 ElseIf Data = "animazione10" Then
  esegui_animazioni_chat_ricevute.Show
  esegui_animazioni_chat_ricevute.Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\chat" & "\immagine10" & ".gif"
 ElseIf Data = "animazione11" Then
  esegui_animazioni_chat_ricevute.Show
  esegui_animazioni_chat_ricevute.Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\chat" & "\immagine11" & ".gif"
  ElseIf Data = "animazione12" Then
   esegui_animazioni_chat_ricevute.Show
  esegui_animazioni_chat_ricevute.Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\chat" & "\immagine12" & ".gif"
 ElseIf Data = "animazione13" Then
  esegui_animazioni_chat_ricevute.Show
  esegui_animazioni_chat_ricevute.Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\chat" & "\immagine13" & ".gif"
 ElseIf Data = "animazione14" Then
  esegui_animazioni_chat_ricevute.Show
  esegui_animazioni_chat_ricevute.Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\chat" & "\immagine14" & ".gif"
 ElseIf Data = "animazione15" Then
  esegui_animazioni_chat_ricevute.Show
  esegui_animazioni_chat_ricevute.Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\chat" & "\immagine15" & ".gif"
 ElseIf Data = "animazione16" Then
  esegui_animazioni_chat_ricevute.Show
  esegui_animazioni_chat_ricevute.Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\chat" & "\immagine16" & ".gif"
 ElseIf Data = "animazione17" Then
  esegui_animazioni_chat_ricevute.Show
  esegui_animazioni_chat_ricevute.Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\chat" & "\immagine17" & ".gif"
 ElseIf Data = "animazione18" Then
  esegui_animazioni_chat_ricevute.Show
  esegui_animazioni_chat_ricevute.Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\chat" & "\immagine18" & ".gif"
 ElseIf Data = "animazione19" Then
  esegui_animazioni_chat_ricevute.Show
  esegui_animazioni_chat_ricevute.Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\chat" & "\immagine19" & ".gif"
 ElseIf Data = "animazione20" Then
  esegui_animazioni_chat_ricevute.Show
  esegui_animazioni_chat_ricevute.Anim1.AnimatedGifPath = App.Path & "\animazioni" & "\chat" & "\immagine20" & ".gif"
 ElseIf Data = "sfondo1" Then
  chat_style.Anim1.AnimatedGifPath = App.Path & "\sfondi" & "\chat" & "\immagine1" & ".jpg"
  MakeTransparent chat.hwnd, 200
 ElseIf Data = "sfondo2" Then
  chat_style.Anim1.AnimatedGifPath = App.Path & "\sfondi" & "\chat" & "\immagine2" & ".jpg"
  MakeTransparent chat.hwnd, 200
 ElseIf Data = "sfondo3" Then
  chat_style.Anim1.AnimatedGifPath = App.Path & "\sfondi" & "\chat" & "\immagine3" & ".jpg"
  MakeTransparent chat.hwnd, 200
 ElseIf Data = "sfondo4" Then
  chat_style.Anim1.AnimatedGifPath = App.Path & "\sfondi" & "\chat" & "\immagine4" & ".jpg"
  MakeTransparent chat.hwnd, 200
 ElseIf Data = "sfondo5" Then
  chat_style.Anim1.AnimatedGifPath = App.Path & "\sfondi" & "\chat" & "\immagine5" & ".jpg"
  MakeTransparent chat.hwnd, 200
 ElseIf Data = "sfondo6" Then
  chat_style.Anim1.AnimatedGifPath = App.Path & "\sfondi" & "\chat" & "\immagine6" & ".jpg"
  MakeTransparent chat.hwnd, 200
 ElseIf Data = "sfondo7" Then
  chat_style.Anim1.AnimatedGifPath = App.Path & "\sfondi" & "\chat" & "\immagine7" & ".jpg"
  MakeTransparent chat.hwnd, 200
 ElseIf Data = "sfondo8" Then
  chat_style.Anim1.AnimatedGifPath = App.Path & "\sfondi" & "\chat" & "\immagine8" & ".jpg"
  MakeTransparent chat.hwnd, 200
 ElseIf Data = "sfondo9" Then
  chat_style.Anim1.AnimatedGifPath = App.Path & "\sfondi" & "\chat" & "\immagine9" & ".jpg"
  MakeTransparent chat.hwnd, 200
 End If
ricevi_comandi_chat.List1.AddItem Data 'add data to listbox
End Sub


Private Sub win_Close()
    For i = 0 To 50
        imwindow(i).Hide
    Next
    frmAdd.Hide
    frmBuddyList.Hide
    frmBuddyList.lstBuddy.Clear
    frmBuddyList.lstOffline.Clear
    timesgotinfo = 0
    frmCreate.Hide
    frmIM.Hide
    frmInfo.Hide
    frmSetInfo.Hide
    
    
    login.Show
    
End Sub

Private Sub win_connect()
 login.Icon = icone.Image2.Picture
 TrayModify Tray_Icon, icone.Image2.Picture
End Sub

Private Sub win_DataArrival(ByVal bytesTotal As Long)
    Dim Buffer As String
    Dim Msg As String
    
    win.GetData Msg, vbString
    Buffer = Msg
    
    nummessages = 0
    If InStr(1, Msg, "\-") > 0 Then
    If Len(Msg) > InStr(1, Msg, "\-") + 1 Then
        For i = 1 To Len(Msg) - 1
            If Mid(Msg, i, 2) = "\-" Then
                nummessages = nummessages + 1
            End If
        Next
    Else
        'msg = Replace(msg, "\-", "")
        nummessages = 1
    End If
    Else
    nummessages = 1
    End If
    For q = 1 To nummessages
        If InStr(1, Buffer, "\-") > 0 Then
            Msg = Mid(Buffer, 1, InStr(1, Buffer, "\-") - 1)
            Buffer = Mid(Buffer, InStr(1, Buffer, "\-") + 2)
        End If
    'If we get a verification on login, request buddylist
    If Msg = "Login Success" Then
        SaveSetting "IMer", "Saved info", "Username", Txtnick
        lblStatus = "Logged in! Downloading Buddy List"
        win.SendData "Send b/l\-"
        
    'The server is sending the client a message....
    ElseIf Left(Msg, 3) = "im-" Then
        pos = InStr(4, Msg, "-")
        sUser = Mid(Msg, 4, (pos + 1) - 5)
        im = Mid(Msg, pos + 1)
        found = False
        countit = 0
        For i = 0 To 50
            If imwindow(i).Caption = "" And countit < 1 Then
                firstfree = i
                countit = 1
            End If
            'These lines check to see if an IM window is already open for the user, if
            'not, it creates a new im window
            If UCase(Left(imwindow(i).Caption, Len(sUser))) = UCase(sUser) Then
                start = Len(imwindow(i).txtHistory.Text)
                imwindow(i).txtHistory.SelStart = Len(imwindow(i).txtHistory.Text)
                imwindow(i).txtHistory.SelText = sUser & ": " & im & vbCrLf
                imwindow(i).txtHistory.SelStart = start
                imwindow(i).txtHistory.SelLength = Len(sUser)
                imwindow(i).txtHistory.SelColor = vbRed
                imwindow(i).txtHistory.SelBold = True
                imwindow(i).txtHistory.SelStart = Len(imwindow(i).txtHistory)
                imwindow(i).txtHistory.SelLength = 1
                FlashWindow imwindow(i).hwnd, 1
                found = True
                If Left(imwindow(i).Caption, Len(sUser)) <> sUser Then
                    imwindow(i).Caption = sUser & " : " & login.Txtnick
                    imwindow(i).Txtnick = sUser
                End If
                i = 21
            End If
        Next

        If found = False Then
            imwindow(firstfree).Caption = sUser & " : " & login.Txtnick.Text
            imwindow(firstfree).Txtnick = sUser
            start = Len(imwindow(firstfree).txtHistory.Text)
            imwindow(firstfree).txtHistory.SelStart = Len(imwindow(firstfree).txtHistory.Text)
            imwindow(firstfree).txtHistory.SelText = sUser & ": " & im & vbCrLf
            imwindow(firstfree).txtHistory.SelStart = start
            imwindow(firstfree).txtHistory.SelLength = Len(sUser)
            imwindow(firstfree).txtHistory.SelColor = vbRed
            imwindow(firstfree).txtHistory.SelBold = True
            imwindow(firstfree).txtHistory.SelStart = Len(imwindow(firstfree).txtHistory)
            imwindow(firstfree).txtHistory.SelLength = 1
            imwindow(firstfree).Show
        End If
    ElseIf Left(Msg, 4) = "msg-" Then
        Message = Mid(Msg, 5)
        MsgBox Message, , "Message"
        If Message = "Error! User doesn't exist!" Or Message = "Error! Wrong password!" Then
            lblStatus = "Wrong Username/Password"
            txtPass = ""
            win.Close
        End If
        If Message = "Error! User already logged in!" Then
            cmdLogin.Caption = "accedi"
            lblStatus = "User already logged in"
            win.Close
        End If
    'Add an online user to the buddylist
    ElseIf Left(Msg, 5) = "cusr-" Then
        frmBuddyList.lstBuddy.AddItem Mid(Msg, 6)
    'Add an offline user to the buddylist
    ElseIf Left(Msg, 5) = "dusr-" Then
        frmBuddyList.lstOffline.AddItem Mid(Msg, 6)
    'Signals end of the buddylist. Open the buddylist
    ElseIf Msg = "End b/l" Then
        login.Hide
        login.Visible = True
        Checkremember.Visible = False
        Label16.Visible = False
        Checkautoconnect.Visible = False
        Label17.Visible = False
        Checkattivasfondi_txtchat.Visible = False
        Label18.Visible = False
        Picture12.Visible = True
        Picture12.Height = 630
        Picture12.Width = 855
        Picture12.Top = 4800
        Picture12.Left = 1920
        animazione_connessione.Timer_animazione_connessione.Enabled = True
        SetParent animazione_connessione.hwnd, login.Picture12.hwnd
        animazione_connessione.Show
        animazione_connessione.Move 0, 0
        Timer_attesa.Enabled = True
        frmBuddyList.Caption = Txtnick & "'s Buddy List"
    ElseIf Left(Msg, 3) = "ss-" Then
        doother = True
        For i = 0 To frmBuddyList.lstBuddy.ListCount
            listbuddy = frmBuddyList.lstBuddy.List(i)
            If UCase(listbuddy) = UCase(Mid(Msg, 4)) Then
                frmBuddyList.lstBuddy.RemoveItem (i)
                frmBuddyList.lstOffline.AddItem listbuddy
                doother = False
            End If
        Next
        
        If doother = True Then
            For i = 0 To frmBuddyList.lstOffline.ListCount
                listbuddy = frmBuddyList.lstOffline.List(i)
                If UCase(listbuddy) = UCase(Mid(Msg, 4)) Then
                    frmBuddyList.lstOffline.RemoveItem (i)
                    frmBuddyList.lstBuddy.AddItem listbuddy
                    'attiviamo il timer che ci anima la icona'
                    icone.Timer1.Enabled = True
                End If
            Next
        End If
    ElseIf Left(Msg, 5) = "ginf-" Then
        pos = InStr(6, Msg, "-")
        sUser = Mid(Msg, 6, pos - 6)
        Info = Mid(Msg, pos + 1)
        If timesgotinfo = 0 Then
            LocalInfo = Info
            timesgotinfo = 1
        Else
            frmInfo.Caption = sUser & "'s Info"
            frmInfo.txtInfo = Info
            frmInfo.Show
        End If
    ElseIf Left(Msg, 6) = "forsn-" Then
        realsn = Mid(Msg, 7)
        If UCase(realsn) = UCase(login.Txtnick) Then
            login.Txtnick = realsn
            frmBuddyList.Caption = realsn & "'s Buddy List"
        Else
            For i = 0 To 50
                If UCase(Left(imwindow(i).Caption, Len(realsn))) = UCase(realsn) Then
                    imwindow(i).Caption = realsn & " : " & login.Txtnick
                    imwindow(i).Txtnick = realsn
                End If
            Next
        End If
    End If
    Next
End Sub


'[Checking The mouse event]
Private Sub form_mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim cEvent As Single
cEvent = X / Screen.TwipsPerPixelX
Select Case cEvent
    Case MouseMove
        Debug.Print "MouseMove"
    Case LeftUp
        Debug.Print "Left Up"
    Case LeftDown
        Debug.Print "LeftDown"
    Case LeftDbClick
        WindowState = 0: Me.Show
    Case MiddleUp
        Debug.Print "MiddleUp"
    Case MiddleDown
        Debug.Print "MiddleDown"
    Case MiddleDbClick
        Debug.Print "MiddleDbClick"
    Case RightUp
        
    Case RightDown
        Debug.Print "RightDown"
    Case RightDbClick
        Debug.Print "RightDbClick"
End Select
  dimensioni_form.Text1.Text = login.Top
  dimensioni_form.Text_login_widh = login.Width
  dimensioni_form.Text_login_height = login.Height
  Cmdprofilo.Font.Underline = False
  Cmdprofilo.MousePointer = vbDefault
  cmdCreate.Font.Underline = False
  cmdCreate.MousePointer = vbDefault
  Label_vediaccount.Font.Underline = False
  Label_vediaccount.MousePointer = vbDefault
  Checkremember.BackColor = &H8000000F
  Checkautoconnect.BackColor = &H8000000F
  Checkattivasfondi_txtchat.BackColor = &H8000000F
  cmdLogin.Style = XP_Button
  Cmdminimizza.ForeColor = &H80000002
  Label20.ForeColor = &H80000002
  If Not Label8.Caption = "" Then
     shape5.Visible = True
  Else
     shape5.Visible = False
  End If
End Sub

 ' questo form e' borderless ( senza bordo), impostiamo la immagine 10'
' come bordo che gli permettera' di muovere il form'
Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage login.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub

Private Sub chameleonButton1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage login.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Cancel = 1
 Cmdexit_Click
End Sub

