Attribute VB_Name = "Module1"
Public textmax As Integer
Public imagemax As Integer
Public addctrl As String


Sub txtadd(x As Single, y As Single)
'run Font Dialog
Dim sFont As SelectedFont
On Error GoTo er
FontDialog.iPointSize = 12 * 10
sFont = ShowFont(frmMain.hWnd, "Times New Roman", True)
' To add labels to frmPage
textmax = textmax + 1
Load frmPage.Label1(textmax)
frmPage.Label1(textmax).Caption = InputBox("Text for the label", "Text Caption")
frmPage.Label1(textmax).Visible = True
frmPage.Label1(textmax).left = x
frmPage.Label1(textmax).top = y
frmPage.Label1(textmax).ForeColor = sFont.lColor
frmPage.Label1(textmax).Font = sFont.sSelectedFont
frmPage.Label1(textmax).Font.Size = sFont.nSize
frmPage.Label1(textmax).ZOrder 0
Exit Sub
er:
MsgBox Err.Description
End Sub

Sub imgadd(x As Single, y As Single)
On Error GoTo er
Dim sOpen As SelectedFile
Dim filename
Dim Count As Integer
'Run Open Dialog
FileDialog.sInitDir = OptDefPath
FileDialog.flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT
FileDialog.sDlgTitle = "Open Image file"
FileDialog.sFilter = "Image files (*.jpg, *.bmp)" & Chr$(0) & "*.jpg;*.bmp"
     sOpen = ShowOpen(frmMain.hWnd)
      If Err.Number <> 32755 And sOpen.bCanceled = False Then
        FileList = sOpen.sLastDirectory
        For Count = 1 To sOpen.nFilesSelected
            FileList = FileList & sOpen.sFiles(Count)
        Next Count
End If
MsgBox FileList
' To add images to frmPage
imagemax = imagemax + 1
Load frmPage.Image1(imagemax)
frmPage.Image1(imagemax).ToolTipText = FileList
frmPage.Image1(imagemax).Visible = True
frmPage.Image1(imagemax).left = x
frmPage.Image1(imagemax).top = y
frmPage.Image1(imagemax).Picture = LoadPicture(FileList)
'Form1.Image1(imagemax).Stretch = True
frmPage.Image1(imagemax).ZOrder 0
Exit Sub
er:
If Err.Number <> 32755 Then
MsgBox Err.Description
End If
End Sub


Sub Savepage(filename As String)
'On Error Resume Next
Dim mv As Object, txt As Object, pic As Object
Dim R, G, B As Long
Dim i As Integer
'Change Back Color to Red Green Blue
R = RGBRed(frmPage.BackColor)
G = RGBGreen(frmPage.BackColor)
B = RGBBlue(frmPage.BackColor)
    
'Create Movie
Set mv = CreateObject("swfobjs.swfMovie")
 With mv
        .SetSize frmPage.Width, frmPage.Height 'Set size
        .SetFrameBkColor R, G, B 'Set back Color
        .SetFrameRate frmNew.txtFPS 'Set Frames per second
 End With
    

For i = 0 To frmPage.Controls.Count - 1
If TypeOf frmPage.Controls(i) Is Label And frmPage.Controls(i).Index <> 0 Then
If frmPage.Controls(i).ToolTipText = "text" Then
'Create Object
 Set txt = CreateObject("swfobjs.swfObject")
    With txt
'Change Fore Color to Red Green Blue
R = RGBRed(frmPage.Controls(i).ForeColor)
G = RGBGreen(frmPage.Controls(i).ForeColor)
B = RGBBlue(frmPage.Controls(i).ForeColor)
'Make text file
        .MakeTextSimple frmPage.Controls(i).Font, frmPage.Controls(i).Caption, frmPage.Controls(i).left, frmPage.Controls(i).top, (Round(frmPage.Controls(i).Font.Size) * 25)
'SetSolidFill R,G,B, Alpha
        .SetSolidFill R, G, B, 255
    End With

mv.AddObject txt 'Add text

End If
End If


'***************************************
'* Still working on this, trying to    *
'* understand the MakePicture function.*
'***************************************

If TypeOf frmPage.Controls(i) Is Image And frmPage.Controls(i).Index <> 0 Then
Set pic = CreateObject("swfobjs.swfObject")
pic.MakePicture 0, 0, (frmPage.Controls(i).Picture.Width), (frmPage.Controls(i).Picture.Height), (frmPage.Controls(i).Picture.Width), (frmPage.Controls(i).Picture.Height), frmPage.Controls(i).ToolTipText
'pic.MakePicture XMin, YMin, XMax, YMax, filename

mv.AddObject pic ' Add picture
End If
Next i

'Create swf
    mv.WriteMovie filename
    
'Clean up
    Set mv = Nothing
    Set txt = Nothing
    Set fpic = Nothing


End Sub


'***********************************
'RGB Convertion codes by Dan Redding
'Got this from the PSC (Color Lab)
'***********************************

Public Function RGBRed(RGBCol As Long) As Integer
If RGBCol = -2147483639 Then 'if form white
RGBRed = 255
Else
'Return the Red component from an RGB Color
    RGBRed = RGBCol And &HFF
End If
End Function

Public Function RGBGreen(RGBCol As Long) As Integer
If RGBCol = -2147483639 Then 'if form white
RGBGreen = 255
Else
'Return the Green component from an RGB Color
    RGBGreen = ((RGBCol And &H100FF00) / &H100)
End If
End Function

Public Function RGBBlue(RGBCol As Long) As Integer
If RGBCol = -2147483639 Then 'if form white
RGBBlue = 255
Else
'Return the Blue component from an RGB Color
    RGBBlue = (RGBCol And &HFF0000) / &H10000
End If
End Function
