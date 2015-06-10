VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPreview 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Report Viewer"
   ClientHeight    =   8055
   ClientLeft      =   120
   ClientTop       =   975
   ClientWidth     =   10935
   HelpContextID   =   115
   Icon            =   "frmPreview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   10935
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   8055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      ExtentX         =   19288
      ExtentY         =   14208
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
      Location        =   ""
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   9240
      ScaleHeight     =   315
      ScaleWidth      =   675
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileCopy 
         Caption         =   "Copy to Clipboard"
         HelpContextID   =   121
      End
      Begin VB.Menu mnuFileBitmap 
         Caption         =   "Save as Bitmap"
         HelpContextID   =   122
      End
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Activate()
WebBrowser1.Move 0, 0, ScaleWidth, ScaleHeight
WebBrowser1.Navigate HTMLFile
frmPreview.Caption = "Report Viewer - " & HTMLFile
End Sub

Private Sub Form_Load()
'Replaced form_load with form_activate since user doesn't have to close form before viewing another file
'WebBrowser1.Move 0, 0, ScaleWidth, ScaleHeight
'WebBrowser1.Navigate HTMLFile
'frmPreview.Caption = "Report Viewer - " & HTMLFile
End Sub
Private Sub Form_Resize()
'resize the webbrowser control to fill the window (ScaleWidth and ScaleHeight)
WebBrowser1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub mnuFileBitmap_Click()
Dim S As String

GetHTMLPic
frmPreview.CommonDialog1.FileName = ""
frmPreview.CommonDialog1.Flags = &H1004
frmPreview.CommonDialog1.DialogTitle = "Save Report as Bitmap Image"
frmPreview.CommonDialog1.Filter = "Bitmap Files (*.bmp)|*.bmp"
frmPreview.CommonDialog1.DefaultExt = "bmp"
frmPreview.CommonDialog1.CancelError = True
frmPreview.CommonDialog1.FilterIndex = 0
frmPreview.CommonDialog1.CancelError = False
frmPreview.CommonDialog1.ShowSave
S = frmPreview.CommonDialog1.FileName
If S <> "" Then
    'if file doesn't have ".bmp" appended (when user types a file name with
    ' a period in it and also doesn't explictly type ".bmp")
    If LCase(Right(S, 4)) <> ".bmp" Then
        S = S + ".bmp"
    End If
    SavePicture Picture1.Picture, S
End If
End Sub

Private Sub mnuFileCopy_Click()
GetHTMLPic
Clipboard.Clear
Clipboard.SetData Picture1.Picture, vbCFBitmap
End Sub
Private Sub GetHTMLPic()
' Capture the html display, minus the gray borders and scroll bars
Set Picture1.Picture = CaptureWindow(Me.hWnd, True, 2, 2, _
            Me.ScaleX(Me.Width, vbTwips, vbPixels) - 30, _
            Me.ScaleY(Me.Height, vbTwips, vbPixels) - 66)
End Sub

