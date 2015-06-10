VERSION 5.00
Begin VB.Form frmVPAScan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VPA Scan Summary"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   105
   Icon            =   "frmVPAScan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   5280
         TabIndex        =   14
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   1680
         TabIndex        =   13
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox txtCase 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         TabIndex        =   1
         Top             =   600
         Width           =   6375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "VPA Input File Selected"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2040
      End
      Begin VB.Label lblFile 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2400
         TabIndex        =   11
         Top             =   240
         Width           =   6375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Case Description"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Start Year"
         Height          =   195
         Index           =   2
         Left            =   2400
         TabIndex        =   9
         Top             =   960
         Width           =   870
      End
      Begin VB.Label lblYear 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "End Year"
         Height          =   195
         Index           =   3
         Left            =   2400
         TabIndex        =   7
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label lblYear 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   6
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Number Ages"
         Height          =   195
         Index           =   4
         Left            =   6000
         TabIndex        =   5
         Top             =   960
         Width           =   1140
      End
      Begin VB.Label lblAges 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   7440
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Number Indices"
         Height          =   195
         Index           =   5
         Left            =   6000
         TabIndex        =   3
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblIndex 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   7440
         TabIndex        =   2
         Top             =   1320
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmVPAScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Close #5
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim S As String
Dim N As Integer

    'check for commas in the case description
    S = txtCase.Text
    N = InStr(S, ",")
    If N > 0 Then
        MsgBox "Case Description Cannot Contain Commas. Please Fix.", vbInformation, "Visual Report Designer"
        Exit Sub
    End If
    ScanVPAResults
End Sub

