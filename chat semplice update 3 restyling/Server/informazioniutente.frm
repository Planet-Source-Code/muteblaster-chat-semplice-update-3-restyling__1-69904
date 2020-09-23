VERSION 5.00
Begin VB.Form informazioniutente 
   BackColor       =   &H80000013&
   Caption         =   "Form1"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3435
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   3435
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frameinformazioni 
      BackColor       =   &H80000013&
      Height          =   6615
      Left            =   600
      TabIndex        =   3
      Top             =   720
      Width           =   2295
      Begin VB.TextBox txtPosizioneSpaziIpUtente 
         Alignment       =   2  'Center
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   3960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPosizioneSpaziAvatar 
         Alignment       =   2  'Center
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
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   3960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtIpUtente 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1680
         Width           =   1935
      End
      Begin VB.PictureBox Picavatar 
         BackColor       =   &H80000013&
         Height          =   1575
         Left            =   240
         ScaleHeight     =   1515
         ScaleWidth      =   1515
         TabIndex        =   10
         Top             =   4680
         Width           =   1575
      End
      Begin VB.TextBox Txtavataramico 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "1"
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox txtRecord 
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtInput 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "  "
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox txtPosizioneSpazi 
         Alignment       =   2  'Center
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
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   3960
         Width           =   615
      End
      Begin VB.TextBox txtip 
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Txtutente 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Labeliputente 
         BackStyle       =   0  'Transparent
         Caption         =   "txtiputente"
         Height          =   255
         Left            =   840
         TabIndex        =   20
         Top             =   1320
         Width           =   735
      End
      Begin VB.Shape Shape1 
         Height          =   1815
         Left            =   120
         Shape           =   5  'Rounded Square
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Label Labelavatar 
         BackStyle       =   0  'Transparent
         Caption         =   "avatar"
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Labelrecord 
         BackStyle       =   0  'Transparent
         Caption         =   "txt record"
         Height          =   255
         Left            =   840
         TabIndex        =   18
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Labelip 
         BackStyle       =   0  'Transparent
         Caption         =   "tx tip"
         Height          =   255
         Left            =   960
         TabIndex        =   17
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Labelimput 
         BackStyle       =   0  'Transparent
         Caption         =   "txt imput"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Labelps 
         BackStyle       =   0  'Transparent
         Caption         =   "txt posizione e spazi"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label Labelutente 
         BackStyle       =   0  'Transparent
         Caption         =   "indice"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   2640
         Width           =   615
      End
   End
   Begin VB.Frame Framenick 
      BackColor       =   &H80000013&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.TextBox Text2 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Labelnick 
         BackStyle       =   0  'Transparent
         Caption         =   "informazioni su"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "informazioniutente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txtIpUtente_Change()
server.txtIpUtente.Text = txtIpUtente.Text
End Sub
