VERSION 5.00
Begin VB.Form Calcolatrice 
   BackColor       =   &H0000C000&
   BorderStyle     =   0  'None
   ClientHeight    =   6015
   ClientLeft      =   3495
   ClientTop       =   2340
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   Begin client.CandyButton Cmdchiudi 
      Height          =   255
      Left            =   3960
      TabIndex        =   19
      Top             =   0
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
   Begin VB.CommandButton CmdNumber 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   2400
      TabIndex        =   18
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton CmdNumber 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   1320
      TabIndex        =   17
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton CmdNumber 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   240
      TabIndex        =   16
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton CmdNumber 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   2400
      TabIndex        =   15
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton CmdNumber 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   1320
      TabIndex        =   14
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton CmdNumber 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   240
      TabIndex        =   13
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton CmdNumber 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   2400
      TabIndex        =   12
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton CmdNumber 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   1320
      TabIndex        =   11
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton CmdNumber 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton CmdNumber 
      BackColor       =   &H8000000D&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   1320
      TabIndex        =   9
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton CmdEqual 
      BackColor       =   &H8000000D&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   7
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton Cmdaddizione 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   6
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton Cmdsottrazione 
      Caption         =   "–"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   5
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Cmddivisione 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   4
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton Cmdmoltiplicazione 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   3
      Top             =   1320
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   3375
      Begin VB.Label LblOutput 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   615
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   3375
      End
   End
   Begin client.chameleonButton chameleonButton1 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   10610
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
      MICON           =   "calcolatrice.frx":0000
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
Attribute VB_Name = "Calcolatrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' con questo programma voglio dimostrare quanto facile e' riuscire a fare un acalcolatrice
'stabiliamo le variabili che ci serviranno semplici regole
Dim Input1 As Variant
Dim Input2 As Variant
Dim Output As Variant
Dim Status1, Status2 As Integer
Dim func As String

Private Sub Cmdchiudi_Click()
 chat.Picture20.Top = 10800
End Sub

' questo comando ci permette di annullare le operazioni senza esser costretti a chiudere il programma
Private Sub CmdClear_Click()
    LblOutput.Caption = ""
    Input1 = ""
    Input2 = ""
    Output = ""
End Sub

Private Sub CmdEqual_Click()
Input2 = LblOutput
' stabiliamo le funzioni che corrispondono alle 4 operazioni
If func = "+" Then LblOutput.Caption = Val(Input1) + Val(Input2)
If func = "-" Then LblOutput.Caption = Val(Input1) - Val(Input2)
If func = "*" Then LblOutput.Caption = Val(Input1) * Val(Input2)
If func = "/" Then
' un occhio di riguardo alla matematica non si puo' dividere per zero
  If Input2 = "0" Then
   LblOutput = "non si puo' dividere per zero"
  Else
   LblOutput.Caption = Val(Input1) / Val(Input2)
  End If
End If
End Sub
' ora stabiliamo i comandi che ci permetteranno di eseguire le 4 operazioni fondamnetali
Private Sub Cmdaddizione_Click()
Input1 = LblOutput
LblOutput = ""
func = "+"
End Sub
' eseguiamo comando di sottrazione
Private Sub Cmdsottrazione_Click()
Input1 = LblOutput
LblOutput = ""
func = "-"
End Sub
'eseguiamo comando di divisione
Private Sub Cmddivisione_Click()
Input1 = LblOutput
LblOutput = ""
func = "/"
End Sub
'eseguiamo comando di moltiplicazione
Private Sub Cmdmoltiplicazione_Click()
Input1 = LblOutput
LblOutput = ""
func = "*"
End Sub
' or poniamo i numeri che ci servono
Private Sub CmdNumber_Click(Index As Integer)
LblOutput = LblOutput & Index
End Sub
