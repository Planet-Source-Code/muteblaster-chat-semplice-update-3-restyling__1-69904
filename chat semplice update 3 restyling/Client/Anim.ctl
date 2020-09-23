VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.UserControl Anim 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1680
   ScaleHeight     =   1680
   ScaleWidth      =   1680
   Begin SHDocVwCtl.WebBrowser WB1 
      Height          =   1320
      Left            =   135
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   1320
      ExtentX         =   2328
      ExtentY         =   2328
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
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
Attribute VB_Name = "Anim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

 
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Enum enBorder
   NONE = 0
   Show = 1
End Enum

Dim m_thisHwnd As Long 'read only
Dim m_thisDC As Long 'read only
Dim twipWid As Long, twipHei As Long

'Default Property Values:
Const m_def_offsetX = 0
Const m_def_offsetY = 0
Const m_def_AnimatedGifPath = ""

'Property Variables:
Dim m_offsetX As Long
Dim m_offsetY As Long
Dim m_AnimatedGifPath As String
 

Private Sub UserControl_Resize()
 
 WB1.Move (m_offsetX - 50), _
          (m_offsetY - 50), _
          (Width - m_offsetX) + 150, _
          (Height - m_offsetY) + 150
  
 Call PrintHtmlToDoc
 
End Sub

'   C:\Documents and Settings\evan.ASTROBRI-47XH2C\Desktop\test.gif
Private Sub UserControl_Show()
 '
 ' get the webbrowsers hwnd and hdc
 Call GetWebHwnd
 ' cause the document_complete event to fire
 WB1.Navigate "about:blank"
 
End Sub

Private Function navImg() As String
  '
  'this functions sizes the gif image
  'based upon the width and height of the usercontrol
  '
  Dim pixwid As Long, pixhei As Long
  pixwid = (Width / Screen.TwipsPerPixelX) - _
           (offsetX / Screen.TwipsPerPixelX) + 5
  pixhei = (Height / Screen.TwipsPerPixelY) - _
           (offsetY / Screen.TwipsPerPixelY) + 5
  
  navImg = _
  "<img border='0' hspace='0' vspace='0' " & _
  "width='" & pixwid & _
  "' height='" & pixhei & "' " & _
  "src='" & m_AnimatedGifPath & "'></body>"
 
End Function
Private Function NavGifHtml() As String
  '
  'create the body tag string which prevents
  'this control from looking or acting like a browser
  '
  NavGifHtml = _
    "<body scroll='no' oncontextmenu='return false' " & _
    "leftmargin='0' rightmargin='0' topmargin='0' " & _
    "bottom='0' marginwidth='0' marginheight='0'>"
  
End Function
Private Sub WB1_DocumentComplete(ByVal pDisp As Object, URL As Variant)

   Call PrintHtmlToDoc
   
End Sub
Private Sub PrintHtmlToDoc()
   
   On Error Resume Next
   Dim wid As Long, hei As Long, rgn As Long
   DoEvents
   WB1.Document.Clear
   WB1.Document.write ""
   WB1.Refresh
   WB1.Document.write NavGifHtml & navImg
   '  in the document complete the browser is
   '  programed to display a 3d border, but for our
   '  purposes it detracts so let cut them out
   wid = (WB1.Width / Screen.TwipsPerPixelX) - 6
   hei = (WB1.Height / Screen.TwipsPerPixelY) - 6
   rgn = CreateRectRgn(3, 3, wid, hei)
   SetWindowRgn m_thisHwnd, rgn, True
   
End Sub

' Find the child window with class name "Shell Embedding".
Private Sub GetWebHwnd()
  Const GW_CHILD As Long = 5
  Const GW_HWNDNEXT As Long = 2
  Dim child_hwnd As Long
  Dim class_name As String * 256

  child_hwnd = GetWindow(hWnd, GW_CHILD)
  Do
      ' See if this is the target class.
      GetClassName child_hwnd, class_name, 256
      If Left$(class_name, Len("Shell Embedding")) = _
          "Shell Embedding" Then
          ' store the hwnd in member var
          m_thisHwnd = child_hwnd
          'lets get the hdc while we are at it
          m_thisDC = GetWindowDC(m_thisHwnd)
          Exit Do
      End If

      ' Get the next child.
      child_hwnd = GetWindow(child_hwnd, GW_HWNDNEXT)
   Loop While child_hwnd <> 0
End Sub


'AnimatedGifPath
Public Property Get AnimatedGifPath() As String
    AnimatedGifPath = m_AnimatedGifPath
End Property
Public Property Let AnimatedGifPath(ByVal New_AnimatedGifPath As String)
    m_AnimatedGifPath = New_AnimatedGifPath
    PropertyChanged "AnimatedGifPath"
    Call UserControl_Resize
End Property
'offsetX
Public Property Get offsetX() As Long
    offsetX = m_offsetX
End Property
Public Property Let offsetX(ByVal New_offsetX As Long)
    m_offsetX = New_offsetX
    PropertyChanged "offsetX"
    Call UserControl_Resize
End Property
'offsetY
Public Property Get offsetY() As Long
    offsetY = m_offsetY
End Property
Public Property Let offsetY(ByVal New_offsetY As Long)
    m_offsetY = New_offsetY
    PropertyChanged "offsetY"
    Call UserControl_Resize
End Property
'ShowBorder
Public Property Get ShowBorder() As enBorder
    ShowBorder = UserControl.BorderStyle
End Property
Public Property Let ShowBorder(ByVal New_ShowBorder As enBorder)
    UserControl.BorderStyle() = New_ShowBorder
    PropertyChanged "ShowBorder"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_AnimatedGifPath = m_def_AnimatedGifPath
    m_offsetX = m_def_offsetX
    m_offsetY = m_def_offsetY
End Sub
'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_AnimatedGifPath = PropBag.ReadProperty("AnimatedGifPath", m_def_AnimatedGifPath)
    m_offsetX = PropBag.ReadProperty("offsetX", m_def_offsetX)
    m_offsetY = PropBag.ReadProperty("offsetY", m_def_offsetY)
    Call UserControl_Resize
    UserControl.BorderStyle = PropBag.ReadProperty("ShowBorder", 1)
End Sub
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("AnimatedGifPath", m_AnimatedGifPath, m_def_AnimatedGifPath)
    Call PropBag.WriteProperty("offsetX", m_offsetX, m_def_offsetX)
    Call PropBag.WriteProperty("offsetY", m_offsetY, m_def_offsetY)
    Call PropBag.WriteProperty("ShowBorder", UserControl.BorderStyle, 1)
End Sub


 

 

