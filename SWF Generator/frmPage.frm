VERSION 5.00
Begin VB.Form frmPage 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleMode       =   0  'User
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   1335
      Index           =   0
      Left            =   1200
      Top             =   2640
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   735
      Index           =   0
      Left            =   840
      TabIndex        =   0
      ToolTipText     =   "text"
      Top             =   720
      Visible         =   0   'False
      Width           =   3135
   End
End
Attribute VB_Name = "frmPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Sets a window to a position on the user's screen
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'Gets the user's cursor coordinates
Private Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'Gets a value from an INI file
Private Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'Writes a value to an INI file
Private Const HWND_TOPMOST = -1
'Makes a Window OnTop of other windows
Private Const HWND_NOTOPMOST = -2
'Makes a Window NotOnTop of other windows
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
'These tells the window to not be moved or resized
Private LastPoint As POINTAPI
'I put this here to subtract the current objects
'X and Y coordinates from LastPoint to make
'The object moveable.
Private TheTracker As Boolean
'If True then let the object be moveable
'If False then not let the object be moveable
Private HoldNumber As Integer
'I put this here to detect the current
'Object thats being manipulated.
Private thepointX As Long
'Holds the X cursor coordinate
Private thepointY As Long
'Holds the Y cursor coordinate
Private HoldButton As String
'Holds the current objects name thats being Manipulated
Private Type POINTAPI
x As Long: y As Long
End Type
'This gets the current X and Y coordinates if
'The user's mouse using the GetCursirPos Delare
Private Const flags = SWP_NOMOVE Or SWP_NOSIZE 'Flags for my StayOnTop Function

Private Function CBMove(TheButton As Object)
'Function that allows you to move objects on runtime.
Dim POINT As POINTAPI
    If TheButton.top < 0 Then TheButton.top = 0: Exit Function
GetCursorPos POINT: thepointX& = (POINT.x - LastPoint.x) * Screen.TwipsPerPixelX: thepointY& = (POINT.y - LastPoint.y) * Screen.TwipsPerPixelY: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheButton.Move TheButton.left + thepointX&, TheButton.top + thepointY&: TheButton.Visible = True
End Function

Private Sub Form_Load()
'Set Project properties
Me.top = 0
Me.left = 0
Me.Width = frmNew.txtWidth
Me.Height = frmNew.txtHeight
Me.BackColor = frmNew.btnColor.BackColor
End Sub

'Allows you to move images on runtime.
Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
 Dim POINT As POINTAPI, z As Integer
    If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then CBMove Image1(Index)
End Sub

'Allows you to move label on runtime.
Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
 Dim POINT As POINTAPI, z As Integer
    If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then CBMove Label1(Index)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'Checks to see what object to add (if any).
If Button = vbLeftButton Then
 If addctrl = "text" Then
        txtadd x, y
        addctrl = "mouse"
        End If
  If addctrl = "image" Then
        imgadd x, y
        addctrl = "mouse"
        End If
        End If
End Sub


