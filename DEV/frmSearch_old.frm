VERSION 5.00
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Data Collection"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8640
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearch_old.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   8640
   Begin VB.Frame Frame1 
      Caption         =   "Enter Search Text"
      Height          =   1575
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   7815
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   2520
         TabIndex        =   2
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtSearch 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   7095
      End
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSearch_Click()
    SearchCollection
End Sub

Private Sub Form_Activate()
    Me.Top = 0
    Me.Left = (Screen.Width - Me.ScaleWidth) \ 2
End Sub

Private Sub Form_Load()
    JList = 0
End Sub
