VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "SWF GENERATOR"
   ClientHeight    =   5385
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8130
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "&Insert"
      Begin VB.Menu mnuText 
         Caption         =   "&Add Text"
      End
      Begin VB.Menu mnuImage 
         Caption         =   "&Add Image"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuBack 
         Caption         =   "&Back Color"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'*****************************************************
'*                  SWF GENERATOR!                   *
'*                                                   *
'*  Hello all, the SWF Generator project is a little *
'* demo I made to check the power of Bukoo, a COM    *
'* object that allows you to create to create Flash  *
'* movies. It is still in Beta testing but is very   *
'* powerful all the same.                            *
'* Most of it is commented but if you need any help  *
'* contact me.                                       *
'* BUGS!: I've only worked with it for a few days and*
'* have not yet been able to completly understand the*
'* 'MakePicture' function. So adding pictures is not *
'* working correctly.                                *
'* Another bug is the sizes, the vb size and the     *
'* Bukoo sizes are not the same...                   *
'* Thanks a lot, hope this helps. please vote!       *
'* Contact me if you have any trouble:               *
'*                                                   *
'* Email: Squash@cv.cl                               *
'* web site:  http://www.SquashProductions.com       *
'*                                                   *
'* Bukoo web site: http://bukoo.sourceforge.net/     *
'*****************************************************

Private Sub MDIForm_Load()
frmNew.Show
End Sub

Private Sub mnuAbout_Click()
MsgBox "By Matthew Woodhams - Squash@cv.cl", vbInformation, "Please Vote!"
End Sub

Private Sub mnuBack_Click()
'Run Color Dialog
Dim sColor As SelectedColor
    sColor = ShowColor(Me.hWnd)
    If sColor.bCanceled = True Then
    Else
    frmPage.BackColor = sColor.oSelectedColor 'Set background to selected color
    End If
End Sub

Private Sub mnuNew_Click()
'Make new project
frmNew.Show
End Sub

Private Sub mnuText_Click()
'Add text to the page.
MsgBox "Click form to add text", vbInformation, "Info"
addctrl = "text" 'used for when user clicks frmPage form.
End Sub

Private Sub mnuImage_Click()
'Add image to the page.
MsgBox "Click form to add image", vbInformation, "Info"
addctrl = "image" 'used for when user clicks frmPage form.
End Sub

Private Sub mnuSave_Click()
'Save shockwave file.
Dim sSave As SelectedFile
On Error Resume Next
FileDialog.sFilter = "ShockWave File (*.swf)" & Chr$(0) & "*.swf"
' See Standard CommonDialog Flags for all options
 FileDialog.flags = OFN_HIDEREADONLY
 FileDialog.sDlgTitle = "SWF save"
 FileDialog.sInitDir = App.Path
 FileDialog.sDefFileExt = "*.swf"
 sSave = ShowSave(Me.hWnd)
Screen.MousePointer = 11 'Give the user an hour glass
Savepage (FileDialog.sFileTitle)
Screen.MousePointer = 0
End Sub
