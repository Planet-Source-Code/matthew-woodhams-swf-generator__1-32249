VERSION 5.00
Begin VB.Form frmNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New File"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   Icon            =   "frmNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   243
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   317
   ShowInTaskbar   =   0   'False
   Tag             =   "a"
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   3120
      Width           =   975
   End
   Begin VB.Frame frGeneral 
      Caption         =   "&General"
      Height          =   2895
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      Begin VB.TextBox txtFPS 
         Height          =   375
         Left            =   1920
         TabIndex        =   8
         Text            =   "2"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtHeight 
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Text            =   "4500"
         Top             =   1000
         Width           =   1095
      End
      Begin VB.TextBox txtWidth 
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Text            =   "4500"
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton btnColor 
         BackColor       =   &H80000009&
         Height          =   375
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Frames per Second:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1635
         Width           =   1455
      End
      Begin VB.Label lbColor 
         BackStyle       =   0  'Transparent
         Caption         =   "Background Color:"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label lbWidth 
         BackStyle       =   0  'Transparent
         Caption         =   "Movie Width:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lbHeight 
         BackStyle       =   0  'Transparent
         Caption         =   "Movie Height:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   3120
      Width           =   975
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
Me.Hide
End Sub

Private Sub btnColor_Click()
'Run Color Dialog
Dim sColor As SelectedColor
    sColor = ShowColor(Me.hWnd)
    If sColor.bCanceled = True Then
    Else
    btnColor.BackColor = sColor.oSelectedColor 'Set background to selected color
    End If
End Sub

Private Sub btnOK_Click()
'Create new Document in MDI Form
Dim frmD As frmPage
Set frmD = New frmPage
frmD.Show
Me.Hide
End Sub

