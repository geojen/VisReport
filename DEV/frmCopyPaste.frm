VERSION 5.00
Begin VB.Form frmCopyPaste 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copy/Paste to Clipboard"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCopyPaste.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5745
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5175
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   1560
         Width           =   4335
      End
      Begin VB.CommandButton cmdPaste 
         Caption         =   "Paste from Clipboard to Highlighted Cells"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   960
         Width           =   4335
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy Highlighted Cells to Clipboard"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmCopyPaste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    ClearGridSelection
End Sub

Private Sub cmdCopy_Click()
    CopyToClipboard
End Sub

Private Sub cmdPaste_Click()
    PasteFromClipboard
End Sub


