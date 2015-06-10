VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmGeneral 
   Caption         =   "General Information"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11985
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGeneral.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8265
   ScaleWidth      =   11985
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   1080
      Picture         =   "frmGeneral.frx":0442
      ScaleHeight     =   2025
      ScaleWidth      =   9705
      TabIndex        =   3
      Top             =   360
      Width           =   9735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Project Log File"
      Height          =   1095
      Left            =   360
      TabIndex        =   1
      Top             =   4440
      Width           =   11175
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Open"
         Height          =   255
         HelpContextID   =   103
         Left            =   9720
         TabIndex        =   7
         ToolTipText     =   "Open an Existing Project"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "Create New"
         Height          =   255
         HelpContextID   =   102
         Left            =   9720
         TabIndex        =   6
         ToolTipText     =   "Create a New Project"
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblFile 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "None Specified"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   9375
      End
   End
   Begin VB.Label lblStartYr 
      AutoSize        =   -1  'True
      Caption         =   "Start Year"
      Enabled         =   0   'False
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   7800
      Width           =   870
   End
   Begin VB.Label lblEndYr 
      AutoSize        =   -1  'True
      Caption         =   "End Year"
      Enabled         =   0   'False
      Height          =   195
      Left            =   1680
      TabIndex        =   4
      Top             =   7800
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Visual Report Designer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Index           =   1
      Left            =   4080
      TabIndex        =   0
      Top             =   3120
      Width           =   3240
   End
End
Attribute VB_Name = "frmGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdNew_Click()
CreateNewLogFile
End Sub

Private Sub cmdOpen_Click()
OpenProjectLog
End Sub


Private Sub Form_Load()
    
    'hide year options
    lblStartYr.Visible = False
    lblEndYr.Visible = False
        
End Sub

