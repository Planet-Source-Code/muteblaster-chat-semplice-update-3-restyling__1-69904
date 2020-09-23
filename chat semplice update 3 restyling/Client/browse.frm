VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form browser 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   2040
      TabIndex        =   1
      Top             =   9480
      Width           =   2175
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   3000
         Left            =   1680
         Top             =   480
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   1680
         Top             =   0
      End
      Begin MSComctlLib.ImageList anim_browser 
         Left            =   0
         Top             =   120
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
               Picture         =   "browse.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "browse.frx":062A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "browse.frx":0C54
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "browse.frx":127E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "browse.frx":18A8
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "browse.frx":1ED2
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "browse.frx":24FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "browse.frx":2B26
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "browse.frx":3150
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "browse.frx":377A
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "browse.frx":3DA4
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "browse.frx":43CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "browse.frx":49F8
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "browse.frx":5022
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "browse.frx":564C
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "browse.frx":5C76
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "browse.frx":62A0
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Image img_anim_browser 
         Height          =   330
         Left            =   840
         Picture         =   "browse.frx":68CA
         Top             =   720
         Width           =   330
      End
   End
   Begin SHDocVwCtl.WebBrowser browse 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      ExtentX         =   6588
      ExtentY         =   8705
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "browser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
 Static frame As Integer
 Timer2.Enabled = True
 frame = frame + 1
 If frame > anim_browser.ListImages.Count Then frame = 1
 img_anim_browser.Picture = anim_browser.ListImages(frame).Picture
 Frame1.Top = 2640
 Frame1.Left = 2040
End Sub

Private Sub Timer2_Timer()
 Frame1.Top = 9480
 Frame1.Left = 2040
 Timer1.Enabled = False
 Timer2.Enabled = False
End Sub
